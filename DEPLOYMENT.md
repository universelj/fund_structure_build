# Deployment Notes

## Production Run
- Prefer a WSGI server instead of the built-in `app.run`.
- Example (4 workers):

```bash
gunicorn -w 4 -b 0.0.0.0:5001 app:app
```

## Optional Environment Variables
- `SECRET_KEY`:固定的密钥，避免重启导致会话失效。
- `SECRET_KEY_FILE`:当未设置 `SECRET_KEY` 时，保存固定密钥的文件路径。
- `OUTPUT_TTL_HOURS`:输出文件的过期小时数。设置为 `0` 或不设置表示不清理。
- `OUTPUT_CLEANUP_INTERVAL_SECONDS`:清理检查的最小间隔秒数，默认 `900`。
- `MAX_XML_BYTES`:保存 XML 的最大字节数，默认 `0`（不限制）。
