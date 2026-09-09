# -*- coding: utf-8 -*-
"""Client kết nối Pancake POS Open API.

Tài liệu: https://docs.pancake.biz/pos/api/
Base URL: https://pos.pages.fm/api/v1
Xác thực: truyền api_key qua query string.
"""
from __future__ import annotations

import time
from typing import Iterator, Optional

import requests

BASE_URL = "https://pos.pages.fm/api/v1"

# Server Pancake thỉnh thoảng trả chậm/lỗi tạm khi kéo trang lớn (1000 đơn)
_RETRY_DELAYS = (5, 15, 30)   # giây chờ trước mỗi lần thử lại


class PancakeError(Exception):
    """Lỗi trả về từ Pancake POS API."""


class PancakeClient:
    def __init__(self, api_key: str, timeout: int = 90):
        self.api_key = api_key
        self.timeout = timeout
        self.session = requests.Session()

    def _get(self, path: str, params: Optional[dict] = None) -> dict:
        """GET (chỉ đọc) - tự thử lại khi timeout / rớt mạng / server lỗi 5xx."""
        params = dict(params or {})
        params["api_key"] = self.api_key
        url = f"{BASE_URL}{path}"
        resp = None
        for delay in (*_RETRY_DELAYS, None):
            try:
                resp = self.session.get(url, params=params, timeout=self.timeout)
            except (requests.Timeout, requests.ConnectionError) as e:
                if delay is None:
                    raise PancakeError(
                        f"Không gọi được {BASE_URL} (mạng chậm hoặc rớt): {e}") from e
                print(f"  Pancake API {type(e).__name__}, thử lại sau {delay}s ...", flush=True)
                time.sleep(delay)
                continue
            if resp.status_code >= 500 and delay is not None:
                print(f"  Pancake API lỗi HTTP {resp.status_code}, "
                      f"thử lại sau {delay}s ...", flush=True)
                time.sleep(delay)
                continue
            break

        if resp.status_code == 401 or resp.status_code == 403:
            raise PancakeError("API key không hợp lệ hoặc không có quyền truy cập (HTTP %s)." % resp.status_code)
        if resp.status_code != 200:
            raise PancakeError(f"Lỗi HTTP {resp.status_code} khi gọi {path}: {resp.text[:300]}")

        try:
            data = resp.json()
        except ValueError as e:
            raise PancakeError(f"Phản hồi không phải JSON: {resp.text[:300]}") from e

        if isinstance(data, dict) and data.get("success") is False:
            raise PancakeError(f"API báo lỗi khi gọi {path}: {data.get('message') or data}")
        return data

    # ------------------------------------------------------------------
    def get_shops(self) -> list[dict]:
        """Danh sách shop mà API key có quyền truy cập."""
        data = self._get("/shops")
        return data.get("shops", [])

    def get_users(self, shop_id: str) -> list[dict]:
        """Danh sách nhân viên của shop (kèm bộ phận trong trường 'department')."""
        data = self._get(f"/shops/{shop_id}/users", {"page_size": 500})
        return data.get("data", [])

    def get_orders_page(
        self,
        shop_id: str,
        start_ts: int,
        end_ts: int,
        page_number: int = 1,
        page_size: int = 100,
        update_status: str = "inserted_at",
        filter_status: Optional[list[int]] = None,
    ) -> dict:
        """Một trang đơn hàng trong khoảng thời gian [start_ts, end_ts] (unix giây).

        filter_status: chỉ lấy đơn đang ở các trạng thái này (vd [4, 5] = đơn hoàn).
        """
        params = {
            "startDateTime": start_ts,
            "endDateTime": end_ts,
            "updateStatus": update_status,  # lọc theo mốc thời gian tạo đơn
            "page_number": page_number,
            "page_size": page_size,
        }
        if filter_status:
            params["filter_status[]"] = list(filter_status)
        return self._get(f"/shops/{shop_id}/orders", params)

    def iter_orders(
        self,
        shop_id: str,
        start_ts: int,
        end_ts: int,
        page_size: int = 100,
        update_status: str = "inserted_at",
    ) -> Iterator[dict]:
        """Duyệt toàn bộ đơn hàng trong khoảng thời gian, tự lật trang."""
        page = 1
        while True:
            data = self.get_orders_page(
                shop_id, start_ts, end_ts,
                page_number=page, page_size=page_size, update_status=update_status,
            )
            orders = data.get("data") or data.get("orders") or []
            for order in orders:
                yield order

            total_pages = data.get("total_pages")
            if total_pages is not None:
                if page >= int(total_pages):
                    break
            elif len(orders) < page_size:
                break
            page += 1
