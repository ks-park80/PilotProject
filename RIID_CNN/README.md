# ☢️ RIID_CNN: 방사선 핵종 식별을 위한 1D CNN 모델

이 프로젝트는 방사선 검출기에서 수집된 스펙트럼 데이터를 분석하여 핵종(예: Cs-137)을 식별하는 **1D Convolutional Neural Network (CNN)** 모델을 구현한 Jupyter Notebook입니다.

## 📌 주요 기능

* **데이터 전처리 및 파싱**:
    * 방사선 검출기 로그 파일(.txt)에서 메타데이터(Count, GC, Temp)와 2048 채널 스펙트럼 추출
    * **엄격한 라인 파싱**: 정규표현식을 사용하여 유효한 데이터만 필터링
    * **섭씨 온도 변환**: Raw Temperature 값을 섭씨(℃)로 변환
    * **자동 라벨링**: 스펙트럼 총 카운트(CPS)를 기준으로 `BKG`(배경방사선)와 `Cs-137` 자동 분류
* **데이터셋 구축 및 저장**:
    * 모델(HH300/SPRD), 센서(CLLBC/CLYC/CSI), 라벨별로 데이터를 그룹화하여 효율적인 `.parquet`, `.npy` 형식으로 저장
* **1D CNN 모델 학습**:
    * **Auxiliary Input**: 스펙트럼뿐만 아니라 온도(Temp), 게인(GC) 정보를 보조 입력으로 활용하여 학습 성능 향상
    * **데이터 증강 및 정규화**: 채널별 Z-score 정규화, Anscombe 변환(Square Root) 적용
    * **클래스 불균형 해소**: 학습 시 Pos/Neg 비율을 고려한 가중치 적용
* **추론 및 평가**:
    * **Tri-class 분류**: `BKG`, `Cs-137`, `Others`(불확실/미학습) 3가지 클래스로 분류
    * **OOD(Out-of-Distribution) 감지**: 학습된 프로토타입과의 마할라노비스 거리를 기반으로 미학습 핵종 거부
    * **물리적 게이트(Peak Gate)**: 에너지 피크 위치를 기반으로 한 추가 검증 로직 적용 가능
* **시각화 도구**:
    * **Interactive Spectrum Browser**: 위젯을 통해 개별 스펙트럼의 Raw/Normalized 형태와 추론 결과를 실시간으로 탐색 가능

## 🛠️ 설치 및 실행 환경

이 프로젝트는 **Google Colab** 또는 **Jupyter Notebook** 환경에서 실행할 수 있습니다.

### 필수 라이브러리
```bash
pip install numpy pandas matplotlib scikit-learn torch

## 🚀 사용 방법

    환경 설정:

        Google Drive를 마운트하고 데이터셋 경로(BASE_DIR)를 설정합니다.

    데이터 전처리:

        build_and_save_datasets() 함수를 실행하여 원본 텍스트 로그를 학습 가능한 형태(.npy, .parquet)로 변환합니다.

    모델 학습:

        train_one_group() 함수를 사용하여 특정 센서/모델 조합에 대한 학습을 수행합니다. 학습된 모델은 processed/models/ 경로에 저장됩니다.

    추론 및 시각화:

        predict_file_auto_routed() 함수로 테스트 파일에 대한 추론을 수행합니다.

        launch_spectrum_browser() 함수를 실행하여 추론 결과를 인터랙티브하게 확인합니다.

## 📂 디렉토리 구조

RIID_CNN/
├── Dataset/          # 원본 로그 파일 (.txt)
├── processed/        # 전처리된 데이터 및 모델 저장소
│   ├── group_*.npy
│   ├── group_*.parquet
│   └── models/       # 학습된 모델 (best.pt, scaler.json)
└── RIID_CNN3.ipynb   # 메인 코드

## 📝 참고 사항

    정규화 방식: chan_z (채널별 Z-score) 방식을 권장합니다.

    모델 아키텍처: CNN1D 클래스는 기본적으로 온도와 게인 정보를 보조 입력(use_aux=True)으로 받도록 설계되었습니다.

    OOD 감지: 프로토타입 거리 기반의 OOD 감지 기능을 통해 학습되지 않은 핵종이나 이상 데이터를 Others로 분류합니다.

## 👨‍💻 작성자

    Author: 박길순