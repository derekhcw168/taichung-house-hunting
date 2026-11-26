# -*- coding: utf-8 -*-
"""Wrapper invoking V2 Objective Third-Party South District Presentation Generator."""
import os
import sys

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPTS_DIR)
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from generate_district_presentations_v2 import build_district_pptx

if __name__ == "__main__":
    build_district_pptx(
        os.path.join(PROJECT_ROOT, "scratch", "south_30yr_master.json"),
        os.path.join(PROJECT_ROOT, "南區30年以上大型社區走勢圖"),
        "00_台中市南區30年以上大型社區_大盤匯總走勢圖.png",
        "台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx",
    )
