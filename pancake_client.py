# -*- coding: utf-8 -*-
"""Client kết nối Pancake POS Open API.

Tài liệu: https://docs.pancake.biz/pos/api/
Base URL: https://pos.pages.fm/api/v1
Xác thực: truyền api_key qua query string.
"""
from __future__ import annotations

from typing import Iterator, Optional

import requests

BASE_URL = "https://pos.pages.fm/api/v1"


class PancakeError(Exception):
    """Lỗi trả về từ Pancake POS API."""


class PancakeClient:
    def __init__(self, api_key: str, timeout: int = 30):
        self.api_key = api_key
        self.timeout = timeout
        self.session = requests.Session()

    def _get(self, path: str, params: Optional[dict] = None) -> dict:
        params = dict(params or {})
        params["api_key"] = self.api_key
        url = f"{BASE_URL}{path}"
        try:
            resp = self.session.get(url, params=params, timeout=self.timeout)
        except requests.ConnectionError as e:
            raise PancakeError(f"Không kết nối được tới {BASE_URL}. Kiểm tra mạng. ({e})") from e

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
    ) -> dict:
        """Một trang đơn hàng trong khoảng thời gian [start_ts, end_ts] (unix giây)."""
        return self._get(
            f"/shops/{shop_id}/orders",
            {
                "startDateTime": start_ts,
                "endDateTime": end_ts,
                "updateStatus": update_status,  # lọc theo mốc thời gian tạo đơn
                "page_number": page_number,
                "page_size": page_size,
            },
        )

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
