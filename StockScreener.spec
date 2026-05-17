# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = [
    'PySide6.QtWebEngineWidgets',
    'PySide6.QtWebEngineCore',
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.formatting.rule',
    'plotly.graph_objects',
    'plotly.subplots',
    'plotly.io',
    'plotly.io.kaleido',
    'yfinance',
    'tradingview_screener',
]
hiddenimports += collect_submodules('tradingview_screener')

# Chỉ exclude third-party không dùng — KHÔNG exclude standard library
excludes = [
    # ML / Deep Learning (>1.5 GB)
    'tensorflow', 'tensorflow_core', 'tensorflow_estimator',
    'keras', 'tensorboard', 'tensorflow_hub',
    'torch', 'torchvision', 'torchaudio',
    'transformers', 'tokenizers', 'huggingface_hub', 'diffusers',
    'onnx', 'onnxruntime',
    'jax', 'jaxlib', 'flax',
    'sklearn', 'scikit_learn',
    'xgboost', 'lightgbm', 'catboost',
    'faiss', 'faiss_cpu',
    'llvmlite', 'numba',

    # Data / Analytics không dùng
    'scipy', 'statsmodels', 'sympy',
    'polars', 'pyarrow',
    'duckdb', 'sqlalchemy',
    'h5py',
    'skimage',

    # Visualization không dùng
    'matplotlib', 'seaborn', 'bokeh', 'altair', 'dash',
    'cv2',

    # Jupyter / IPython
    'IPython', 'jupyter', 'notebook', 'jupyterlab',
    'ipykernel', 'ipywidgets', 'nbformat', 'nbconvert',

    # GUI frameworks khác
    'tkinter', 'wx',

    # Cloud / không dùng
    'boto3', 'botocore', 'azure',
    'googleapiclient',

    # Misc
    'docutils', 'sphinx',
    'scrapy',
    'grpc',
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
