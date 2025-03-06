import streamlit as st
import pandas as pd
import numpy as np
import io

def process_vehicle_data(df, vehicle_numbers, month):
    # 기존 코드와 동일하나 파일 경로 대신 바이트 버퍼 사용
    output_buffer = io.BytesIO()
    writer = pd.ExcelWriter(output_buffer, engine='openpyxl')
    
    # 각 차량별 데이터 처리
    vehicle_totals = {}
    for vehicle_num in vehicle_numbers:
        # 특정 차량 데이터 추출
        vehicle_data = df[df['차량번호'] == vehicle_num].copy()

        # 필요한 컬럼 선택 및 이름 변경
        columns_to_select = {
            '날짜': '날짜',
            '차량번호': '차량번호',
            '상차지': '상차지',
            '하차지': '하차지',
            '품목': '품목',
            '횟수': '횟수',
            '수량': '수량',
            '지급단가.1': '지급단가',
            '지급운반비': '지급운반비',
            '※ 특이사항': '※ 특이사항'
        }

        # 선택한 컬럼만 유지
        result_df = vehicle_data[list(columns_to_select.keys())].copy()
        result_df.rename(columns=columns_to_select, inplace=True)

        # 날짜 형식 변경
        result_df['날짜'] = pd.to_datetime(result_df['날짜']).dt.strftime('%Y.%m.%d')

        # 합계 행 계산
        trip_sum = result_df['횟수'].sum() if '횟수' in result_df.columns else 0
        quantity_sum = result_df['수량'].sum() if '수량' in result_df.columns else 0
        fee_sum = result_df['지급운반비'].sum() if '지급운반비' in result_df.columns else 0

        # 합계 행 추가
        total_row = pd.DataFrame({
            '날짜': ['합계'],
            '차량번호': [vehicle_num],
            '상차지': [''],
            '하차지': [''],
            '품목': [''],
            '횟수': [trip_sum],
            '수량': [quantity_sum],
            '지급단가': [''],
            '지급운반비': [fee_sum],
            '※ 특이사항': ['']
        })

        # 원본 데이터와 합계 행 결합
        final_df = pd.concat([result_df, total_row], ignore_index=True)

        # 각 차량별 시트에 저장
        final_df.to_excel(writer, sheet_name=str(vehicle_num), index=False)

        # 각 차량의 총 지급운반비 저장
        vehicle_totals[vehicle_num] = fee_sum

    # 내부/지입 요약 시트
    internal_total = vehicle_totals.get(5154, 0)  # 5154가 없을 경우 0으로 처리
    leased_total = sum(total for vehicle, total in vehicle_totals.items() if vehicle != 5154)

    summary_df = pd.DataFrame([
        {'구분': '내부(5154)', '총 지급운반비': internal_total},
        {'구분': '지입(5154 제외)', '총 지급운반비': leased_total}
    ])
    summary_df.to_excel(writer, sheet_name='내부,지입', index=False)
    
    writer.close()
    output_buffer.seek(0)
    return output_buffer

def main():
    st.title('운반내역 분석기')
    
    # 월 선택
    month = st.selectbox('몇 월 데이터를 분석하시겠습니까?', 
                         options=list(range(1, 13)),
                         format_func=lambda x: f'{x}월')
    
    # 파일 업로드
    uploaded_file = st.file_uploader(f"{month}월 엑셀 파일을 업로드해주세요", 
                                   type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(uploaded_file, sheet_name=f"{month}월", header=1)
            
            # 데이터프레임 미리보기
            st.subheader('데이터 미리보기')
            st.dataframe(df.head())
            st.write(f'총 {len(df)} 행의 데이터가 있습니다.')
            
            # 차량 번호 목록
            vehicle_numbers = [5154, 5366, 5411, 7051, 7180, 7419, 7475, 7843, 8599, 9210]
            
            # 실제 존재하는 차량 번호만 필터링
            existing_vehicles = [v for v in vehicle_numbers if v in df['차량번호'].unique()]
            if len(existing_vehicles) < len(vehicle_numbers):
                missing = set(vehicle_numbers) - set(existing_vehicles)
                st.warning(f"주의: 다음 차량 번호는 데이터에 존재하지 않습니다: {missing}")
            
            # 분석 실행 버튼
            if st.button('분석 실행'):
                with st.spinner('데이터 처리 중...'):
                    # 데이터 처리
                    output_buffer = process_vehicle_data(df, existing_vehicles, month)
                    
                    # 파일 다운로드 버튼 생성
                    st.success('분석이 완료되었습니다!')
                    st.download_button(
                        label='결과 엑셀 파일 다운로드',
                        data=output_buffer,
                        file_name=f"{month}월_운반내역_차량별.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
        except Exception as e:
            st.error(f'오류 발생: {e}')
            st.error('파일 형식이 올바른지 확인해주세요.')

if __name__ == '__main__':
    main()