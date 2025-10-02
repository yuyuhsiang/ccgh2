from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os

app = Flask(__name__)
DATA_DIR = 'data'
EXCEL_FILE = os.path.join(DATA_DIR, 'e.xlsx')
CSV_FILE   = os.path.join(DATA_DIR, 'dialysis_data.csv')

def prepare_csv():
    """讀取 Excel 並儲存為 CSV，保留 ROC 年-月格式。"""
    df = pd.read_excel(EXCEL_FILE, header=None)
    df_subset = df.iloc[0:4, :]
    df_subset.columns = df_subset.iloc[0]
    df_data = df_subset.iloc[1:4]
    df_t = df_data.T.reset_index()
    df_t.columns = ['year_month', 'hemo_visits', 'peri_visits', 'hemo_patients']
    df_t = df_t[df_t['year_month'] != '年份/月份']
    df_t = df_t.dropna(subset=['hemo_visits']).reset_index(drop=True)
    df_t['hemo_visits']   = pd.to_numeric(df_t['hemo_visits'], errors='coerce')
    df_t['peri_visits']   = pd.to_numeric(df_t['peri_visits'], errors='coerce')
    df_t['hemo_patients'] = pd.to_numeric(df_t['hemo_patients'], errors='coerce')
    df_t['peri_ratio'] = df_t['peri_visits'] / (df_t['peri_visits'] + df_t['hemo_visits'])
    os.makedirs(DATA_DIR, exist_ok=True)
    df_t.to_csv(CSV_FILE, index=False)

prepare_csv()
df = pd.read_csv(CSV_FILE, dtype={'year_month': str})

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/data')
def api_data():
    start = request.args.get('start', '110-01')
    end   = request.args.get('end',   '114-12')
    mask = (df['year_month'] >= start) & (df['year_month'] <= end)
    data = df.loc[mask].sort_values('year_month')
    return jsonify({
        'labels':       data['year_month'].tolist(),
        'hemo_visits':  data['hemo_visits'].tolist(),
        'peri_visits':  data['peri_visits'].tolist(),
        'hemo_patients':data['hemo_patients'].tolist(),
        'peri_ratio':   data['peri_ratio'].round(3).tolist()
    })

@app.route('/api/summary')
def api_summary():
    total = 4
    completed = 4
    abnormal = 0
    return jsonify({
        'total_metrics':     total,
        'completed_metrics': completed,
        'abnormal_rate':     abnormal,
        'completion_rate':   round(completed/total*100)
    })

if __name__ == '__main__':
    app.run(debug=True)
