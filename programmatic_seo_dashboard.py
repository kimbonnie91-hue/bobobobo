"""[통합됨] 경쟁사 분석은 lf_seo_dashboard.py 의 '경쟁사 분석' 뷰로 이동했습니다.

하위 호환을 위해 이 파일은 통합 대시보드를 실행합니다.
실행 권장: streamlit run lf_seo_dashboard.py
"""
import runpy
import os

runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "lf_seo_dashboard.py"), run_name="__main__")
