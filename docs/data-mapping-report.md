# Verified data-mapping report

Generated during initial implementation on 2026-10-21.

No OMS, Master NDR, or courier MIS workbooks were present in the supplied workspace. Therefore no source column mapping can be marked verified. The engine uses the alias configuration in `config/default.yaml`, validates headers before import, and rejects files when required fields cannot be resolved. Replace or extend those aliases after uploading representative source files; do not treat the defaults as verified source truth.

Required canonical fields: `warehouse`, `order_no`, `docket_no`, `order_date`, `shipping_date`, `customer_name`, `mobile`, `alternate`, `shipment`, `pincode_zone`, `current_status`.

Courier-specific mappings are intentionally not guessed. Courier MIS files must include `docket_no` plus any of `pdd`, `courier_remark`, `courier_attempts`, `current_status`, and `latest_scan_timestamp`, or be mapped in configuration.
