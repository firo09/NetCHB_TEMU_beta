from flask import Flask, send_from_directory
import os, json

app = Flask(__name__, static_folder='static')

# --- 新增：注册 API 蓝图 ---
from validator_api import bp as validator_bp
app.register_blueprint(validator_bp)

# --- 新增：CORS（仅 /api/* 开放到白名单） ---
try:
    from flask_cors import CORS
    # 环境变量 CORS_ORIGINS 可以是 JSON 列表或用逗号分隔的字符串
    raw = os.environ.get("CORS_ORIGINS", "")
    try:
        origins = json.loads(raw) if raw.strip().startswith("[") else [o.strip() for o in raw.split(",") if o.strip()]
    except Exception:
        origins = [o.strip() for o in raw.split(",") if o.strip()]
    if not origins:
        origins = []  # 没给就不开放
    CORS(app, resources={r"/api/*": {"origins": origins}})
except Exception:
    pass

# 根路径返回静态 index.html
@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

# 其它静态资源（.js/.css/.json/.png …）
@app.route('/<path:path>')
def static_proxy(path):
    return send_from_directory('static', path)

# Heroku 会读取 PORT 环境变量
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
