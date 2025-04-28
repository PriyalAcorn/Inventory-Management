import pandas as pd
import pyodbc

# SQL Server connection
conn = pyodbc.connect(
    'DRIVER={SQL Server};'
    'SERVER=bcdwsqlserver.database.windows.net;'
    'DATABASE=BCDWsqldatabase;'
    'UID=bcsqlsvr;'
    'PWD=Ac0rn$ql188' 
)

# SQL query
query = """
SELECT 
    a.[Item_no],
    a.[Remaining_QTY],
    a.[Name],
    b.[PK-Size],
    ROUND(a.[Remaining_QTY] / CAST(b.[PK-Size] AS FLOAT), 0) AS NoOfPacks
FROM 
    Inventory_ASL a
INNER JOIN 
    [dbo].[Acorn Solutions Ltd_ GB$Item$437dbf0e-84ff-417a-965d-ed2bb9650972] b 
    ON a.[Item_no] = b.[Single Item Code]
WHERE 
    ISNUMERIC(b.[PK-Size]) = 1
    AND TRY_CAST(b.[PK-Size] AS FLOAT) IS NOT NULL
    AND TRY_CAST(b.[PK-Size] AS FLOAT) <> 0;
"""

# Run the query
df = pd.read_sql_query(query, conn)

# Network path to save Excel file
save_path = r'\\172.30.36.151\Amazon Data\Acorn Solution Limited\Inventory\inventory_output.xlsx'

# Save to Excel
df.to_excel(save_path, index=False)

print(f"Excel file saved successfully at {save_path}")
