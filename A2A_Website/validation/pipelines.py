import pandas as pd
import json
import os
from .steps import *  # step 함수 다 임포트

PIPELINE_PATH = os.path.join(os.path.dirname(__file__), "user_pipelines.json")

def load_country_pipelines():
    if not os.path.exists(PIPELINE_PATH):
        return {}
    with open(PIPELINE_PATH, encoding="utf-8") as f:
        return json.load(f)

def save_country_pipeline(country, steps):
    pipelines = load_country_pipelines()
    pipelines[country] = steps
    with open(PIPELINE_PATH, "w", encoding="utf-8") as f:
        json.dump(pipelines, f, ensure_ascii=False, indent=2)

def get_all_step_functions():
    from .steps import list_all_step_functions
    return list_all_step_functions()

def run_country_validation(country, input_path, hawb_path, target_year, target_month, output_folder):
    pipelines = load_country_pipelines()
    steps = pipelines.get(country)
    if not steps:
        raise ValueError(f"No pipeline defined for country: {country}")
    # 첫 step은 input_path/target_year/target_month 필요, 이후는 DataFrame
    df_or_path = globals()[steps[0]](input_path, target_year, target_month, hawb_path=hawb_path, output_folder=output_folder)
    for step in steps[1:]:
        fn = globals()[step]
        if isinstance(df_or_path, pd.DataFrame):
            df_or_path = fn(
                df_or_path,
                hawb_path=hawb_path,
                year=target_year,
                month=target_month,
                output_folder=output_folder
            )
        else:
            break
    return df_or_path
# 