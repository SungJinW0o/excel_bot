import os
import pandas as pd

os.makedirs('input_data', exist_ok=True)

rows = [
    [1001, 'Acme Corp', 'Widget A', 'Widgets', 10, 25.0, '2026-02-05', 'Completed', 'North'],
    [1002, 'Beta LLC', 'Widget B', 'Widgets', 0, 20.0, '2026-02-05', 'Completed', 'South'],
    [1003, 'Cyan Inc', 'Gadget X', 'Gadgets', 5, 0.0, '2026-02-05', 'Completed', 'East'],
    [1004, 'Delta Co', 'Gadget Y', 'Gadgets', 3, 40.0, '2026-02-05', 'Cancelled', 'West'],
    [1005, 'Echo Ltd', 'Service Plan', 'Services', 2, 150.0, '2026-02-06', 'Pending', 'North'],
    [1006, 'Foxtrot GmbH', 'Widget A', 'Widgets', 7, 25.0, '2026-02-06', 'Completed', 'North'],
    [1007, 'Gamma SA', 'Gadget X', 'Gadgets', 4, 45.0, '2026-02-06', 'Completed', 'South'],
    [1008, 'Helix PLC', 'Service Plan', 'Services', 1, 150.0, '2026-02-06', 'Completed', 'West'],
]

columns = [
    'OrderID', 'ClientName', 'Product', 'Category', 'Quantity',
    'UnitPrice', 'OrderDate', 'Status', 'Region'
]

df = pd.DataFrame(rows, columns=columns)
df['OrderDate'] = pd.to_datetime(df['OrderDate'])

output_path = os.path.join('input_data', 'sales_data_2026-02-06.xlsx')

df.to_excel(output_path, index=False)
print(f'Created {output_path}')
