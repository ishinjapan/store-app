import os
import runpy

os.environ["STORE_APP_VARIANT"] = "store"

runpy.run_module("store_app", run_name="__main__")
