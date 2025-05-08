import os
import pandas as pd


def convert_excel(input_path):
    df = pd.read_excel(input_path)

    # 변환 로직 예시: 컬럼명을 대문자로 변환
    df.columns = [col.upper() for col in df.columns]

    output_dir = os.path.dirname(input_path)
    output_path = os.path.join(output_dir, "converted_output.xlsx")

    df.to_excel(output_path, index=False)
    return output_path
