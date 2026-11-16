# Verified Master NDR mapping

Source inspected: `OCT - MONTH NDR.xlsx`, supplied from `D:\Rishi_Gowtham_file\Desktop\ndr records\`.

| Field | Column | Row 1 | Row 2 |
|---|---:|---|---|
| Warehouse | A | Warehouse | Warehouse |
| OrderNo | B | OrderNo | OrderNo |
| Order Date | C | Order Date | Order Date |
| Customer | D | Cx Name | Cx Name |
| Mobile | G | Mobile No | Mobile No |
| Alternate | H | Alt No | Alt No |
| Shipping date | I | ShippingDate | ShippingDate |
| Shipment | J | Shipment | Shipment |
| Pincode zone | K | blank | PincodeZone |
| Docket | L | Docketno | Docketno |
| Our EDD | M | OUR EDD | OUR EDD |
| PDD | N | blank | PDD |
| OMS status | O | OMS STATUS | OMS STATUS |
| Current status | P | blank | Current status |
| Latest courier remark | Q | blank | Courier remark |
| Daily group | R:T, U:W, ... | date in Row 1 | Agent/remarks fields in Row 2 |
| Global attempts | AY | Attempts | Attempts |

Daily groups are three columns: Courier remark, Agent for NDR, and Our Remarks. The next date group is inserted immediately before the global Attempts column. Historical groups are never removed.

Observed on 2026-10-21: 6,696 data rows; 4,110 Delhivery rows; 3,307 rows with non-zero global Attempts. Applying the supplied eligibility rules produced an approximate dry-run of 75 Delhivery candidates and 244 Local candidates. These counts are for workbook inspection only until the corresponding courier MIS and previous-day calling files are supplied.
