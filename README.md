
```markdown
# Expenses Management with Excel

This is a Django application for managing expenses with Excel. You can make, update, and retrieve Excel data records using the following API endpoints:

## API Endpoints

### make Excel Sheet

Endpoint: `POST /make_excel/<str:sheet_name>/`

make a new Excel sheet with the specified `sheet_name`. This API expects the following request body:

```json
{
    "json_objects": [
        {
            "category": "F",  // Use "F" for farmer
            "id": 101,
            "description": "Apply for expenses unknown unknown expenses unknown for testing purpose",
            "payment_mode": "upi",
            "bank": "HDFC",
            "amount": 5000.0,
            "complaint": "This is the complaint raised by the farmer"
        },
        // Add more objects as needed
    ]
}
```

### make Daily Summary Sheet

Endpoint: `POST /make_daily_summary_sheet/<str:sheet_name>/`

make a daily summary sheet with the specified `sheet_name`. This API expects the following request body:

```json
{}
```

You can customize the request bodies and endpoints as needed for your application. Make sure to specify the format of the data expected in the request bodies and include any necessary data validation.

### Expense Category Codes

- "F" for farmer
- "V" for vehicles
- "S" for shops
- "O" for other expenses
- "W" for vouchers

Feel free to adapt this README to suit your project's requirements and include any additional information or installation instructions if necessary.
```