import pandas as pd
import json
import re
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse

app = FastAPI()

@app.post("/uploadfile/")
async def upload_excel_file(file: UploadFile):
    # 检查文件类型是否为Excel
    if file.filename.endswith('.xlsx'):
        try:
            # 保存上传的Excel文件到服务器
            with open(file.filename, "wb") as f:
                f.write(file.file.read())

            # 读取上传的Excel文件
            df = pd.read_excel(file.filename)

            # 根据字段A和字段B进行分组，并对数值列进行求和

            df_grouped = df.groupby(['判断用字段A', '判断用字段B'])['数值'].sum()
            print(df_grouped)
            
            
            # 将合并后的结果填充到原始DataFrame中的新列中
            df['合并结果'] = df.apply(lambda row: df_grouped[(df_grouped['判断用字段A'] == row['判断用字段A']) & (df_grouped['判断用字段B'] == row['判断用字段B'])]['数值'].values[0] if pd.notna(row['判断用字段A']) else None, axis=1)
            print(df)
            # 将数值列转换为字符串
            df['数值'] = df['数值'].apply(lambda x: str(x) if not pd.isna(x) else None)



            return JSONResponse(content={"message": "File uploaded and processed successfully", "data": '1'})
        except Exception as e:
            return JSONResponse(content={"message": "Error processing file", "error": str(e)}, status_code=500)
    else:
        return JSONResponse(content={"message": "Invalid file format. Only Excel files (.xlsx) are allowed."}, status_code=400)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
