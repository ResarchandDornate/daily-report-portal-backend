"""MinIO client wrapper for sales upload sheets.

Reads connection info from env vars (see .env):
  MINIO_ENDPOINT, MINIO_ACCESS_KEY, MINIO_SECRET_KEY, MINIO_BUCKET, MINIO_SECURE

Lazy-initialised — first call ensures the bucket exists.  Subsequent calls
reuse the same client.
"""
from __future__ import annotations

import io
import os
from datetime import timedelta
from functools import lru_cache

from minio import Minio
from minio.error import S3Error


@lru_cache(maxsize=1)
def _client() -> Minio:
    endpoint = os.environ.get("MINIO_ENDPOINT", "minio:9000")
    access_key = os.environ.get("MINIO_ACCESS_KEY") or os.environ.get(
        "MINIO_ROOT_USER", "minioadmin"
    )
    secret_key = os.environ.get("MINIO_SECRET_KEY") or os.environ.get(
        "MINIO_ROOT_PASSWORD", "minioadmin"
    )
    secure = os.environ.get("MINIO_SECURE", "false").lower() == "true"
    return Minio(endpoint, access_key=access_key, secret_key=secret_key, secure=secure)


def bucket_name() -> str:
    return os.environ.get("MINIO_BUCKET", "sales-uploads")


def _ensure_bucket(client: Minio, bucket: str) -> None:
    try:
        if not client.bucket_exists(bucket):
            client.make_bucket(bucket)
    except S3Error:
        # Race condition or already exists — safe to ignore.
        pass


def put_object(
    object_key: str,
    data: bytes,
    content_type: str = "application/octet-stream",
) -> None:
    """Upload `data` bytes to MinIO under `object_key`."""
    client = _client()
    bucket = bucket_name()
    _ensure_bucket(client, bucket)
    client.put_object(
        bucket,
        object_key,
        io.BytesIO(data),
        length=len(data),
        content_type=content_type,
    )


def get_object_bytes(object_key: str) -> bytes:
    """Read the full file bytes from MinIO."""
    client = _client()
    bucket = bucket_name()
    response = client.get_object(bucket, object_key)
    try:
        return response.read()
    finally:
        response.close()
        response.release_conn()


def delete_object(object_key: str) -> None:
    """Best-effort delete — swallow not-found so callers can be idempotent."""
    client = _client()
    bucket = bucket_name()
    try:
        client.remove_object(bucket, object_key)
    except S3Error:
        pass


def presigned_download_url(
    object_key: str,
    expires_minutes: int = 10,
    download_filename: str | None = None,
) -> str:
    """Return a short-lived presigned URL the browser can hit directly.

    `download_filename` triggers Content-Disposition: attachment, so the
    browser downloads the file with the original name rather than the
    UUID-prefixed object key.
    """
    client = _client()
    bucket = bucket_name()
    _ensure_bucket(client, bucket)
    extra_headers = {}
    if download_filename:
        extra_headers["response-content-disposition"] = (
            f'attachment; filename="{download_filename}"'
        )
    return client.presigned_get_object(
        bucket,
        object_key,
        expires=timedelta(minutes=expires_minutes),
        response_headers=extra_headers or None,
    )
