# Excel Data Analysis Assignment

Basic data analysis using Microsoft Excel on a product dataset.

### Functions Used

| Function | Formula | Purpose |
|---|---|---|
| SUM | `=SUM(D2:D35)` | Total price |
| COUNT | `=COUNT(D2:D35)` | Count products |
| AVERAGE | `=AVERAGE(D2:D35)` | Average price |
| MIN | `=MIN(D2:D35)` | Minimum price |
| MAX | `=MAX(D2:D35)` | Maximum price |
| IF | `=IF(D2>=500,"High Price","Standard Price")` | Price categorization |
| SUMIF | `=SUMIF(F2:F35,"Electronics",D2:D35)` | Electronics total |
| COUNTIF | `=COUNTIF(D2:D35,"<100")` | Products below $100 |
| LEFT | `=LEFT(A2,2)` | Extract Day |
| RIGHT | `=RIGHT(A2,2)` | Extract Country Code |
| MID | `=MID(A2,4,3)` | Extract Month |

### Results

- Total Products: 34
- Total Price: $10,100
- Average Price: $297.06
- Minimum Price: $30
- Maximum Price: $1,000
- Electronics Total: $8,050
- Products Below $100: 11

### Tool
Microsoft Excel
