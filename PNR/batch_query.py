import sys
sys.path.insert(0, 'c:/Users/Administrator/OneDrive/桌面')
from pnr_extractor import process_excel_file

# 处理当前目录的PNR批量查询文件
input_file = r'c:\Users\Administrator\OneDrive\桌面\MU,8L大客户打标脚本\PNR\pnr_batch_query.xlsx'
output_file = r'c:\Users\Administrator\OneDrive\桌面\MU,8L大客户打标脚本\PNR\pnr_batch_query_result2.xlsx'

print(f"输入文件: {input_file}")
print(f"输出文件: {output_file}")
print("=" * 60)

success = process_excel_file(input_file, output_file, "KMG319")

if success:
    print("\n[OK] Processing completed!")
else:
    print("\n[ERROR] Processing failed!")
