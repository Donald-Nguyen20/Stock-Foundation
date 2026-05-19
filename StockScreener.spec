# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = [
    # PySide6 WebEngine (dùng để hiển thị chart plotly)
    'PySide6.QtWebEngineWidgets',
    'PySide6.QtWebEngineCore',
    # openpyxl — import động bên trong hàm _dashboard_sheet
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.formatting.rule',
    'openpyxl.chart',
    'openpyxl.chart.series',
    # plotly / kaleido
    'plotly.graph_objects',
    'plotly.subplots',
    'plotly.io',
    'plotly.io.kaleido',
    # yfinance & screener
    'yfinance',
    'tradingview_screener',
]
hiddenimports += collect_submodules('tradingview_screener')
hiddenimports += collect_submodules('openpyxl')

excludes = [
    # ── ML / Deep Learning (>1.5 GB) ──────────────────────────────
    'tensorflow', 'tensorflow_core', 'tensorflow_estimator',
    'keras', 'tensorboard', 'tensorflow_hub', 'tf_keras',
    'torch', 'torchvision', 'torchaudio',
    'transformers', 'tokenizers', 'huggingface_hub', 'diffusers',
    'sentence_transformers', 'safetensors',
    'onnx', 'onnxruntime',
    'jax', 'jaxlib', 'flax',
    'sklearn', 'scikit_learn',
    'xgboost', 'lightgbm', 'catboost',
    'faiss', 'faiss_cpu',
    'llvmlite', 'numba',

    # ── Data / Analytics ──────────────────────────────────────────
    'scipy', 'statsmodels', 'sympy',
    'polars', 'pyarrow',
    'duckdb', 'sqlalchemy',
    'h5py',
    'skimage', 'scikit_image',
    'ydata_profiling', 'visions',
    'shapely',

    # ── Visualization không dùng ──────────────────────────────────
    'matplotlib', 'seaborn', 'bokeh', 'altair', 'dash',
    'cv2', 'imageio', 'tifffile', 'pywavelets',
    'wordcloud', 'squarify',

    # ── Excel alternatives (chỉ dùng openpyxl) ────────────────────
    'xlsxwriter', 'xlrd', 'xlwt', 'pyxlsb',

    # ── Document / PDF ────────────────────────────────────────────
    'reportlab', 'fpdf2',
    'docx', 'python_docx',
    'python_pptx', 'pptx',

    # ── OCR / Image processing ────────────────────────────────────
    'pytesseract', 'rapidocr_onnxruntime',

    # ── Web frameworks ────────────────────────────────────────────
    'fastapi', 'uvicorn', 'starlette', 'choreographer',
    'scrapy',

    # ── Jupyter / IPython ─────────────────────────────────────────
    'IPython', 'jupyter', 'notebook', 'jupyterlab',
    'ipykernel', 'ipywidgets', 'nbformat', 'nbconvert',

    # ── GUI frameworks khác ───────────────────────────────────────
    'tkinter', 'wx',

    # ── Vietnamese stock libs (không dùng trực tiếp) ──────────────
    'vnstock', 'vnstock3', 'vnstock_ezchart', 'vnai',

    # ── Cloud / không dùng ────────────────────────────────────────
    'boto3', 'botocore', 'azure',
    'googleapiclient',

    # ── Dev / Docs ────────────────────────────────────────────────
    'pytest', 'pytest_timeout',
    'docutils', 'sphinx',
    'grpc', 'grpcio',

    # ── Scientific misc ───────────────────────────────────────────
    'coolprop', 'iapws',
]

a = Analysis(
    ['screener_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('python screen_top100.py', '.'),
        ('image.ico', '.'),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='StockScreener',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['image.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='StockScreener',
)
