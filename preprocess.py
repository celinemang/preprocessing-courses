import os
import re
import pandas as pd



def process_excel(file_path, output_dir):
    # 엑셀 파일 로드
    df = pd.read_excel(file_path)
    

    # 특정 열 제거

    df.drop(columns=[
        'Subjects', 
        'Numbers', 
        'Sections', 
        'Abbreviated Titles',  
        'Level Restrictions',
        'Campus Restrictions',
        'Professor Ratings',

        ],axis = 1, inplace=True)


    #TBA 제거
    df['Days'] = df['Days'].apply(lambda x: x.replace('TBA', '') if pd.notna(x) else x)
    df['Times'] = df['Times'].apply(lambda x: '' if x == 'TBA-TBA' else x)
    df['Credits'] = df['Credits'].apply(lambda x: '' if x == 'TBA' else x)

    # 시간 정보를 추출하기 위한 정규 표현식 패턴
    pattern = r'\d{1,2}:\d{2}[ap]m-\d{1,2}:\d{2}[ap]m'
    def extract_time(text):
        match = re.search(pattern, text)
        if match:
            return match.group()  # 매칭된 시간 정보 반환
        else:
            return ''  # 매칭되는 시간 정보가 없으면 None 반환

    # 'Times' 칼럼에 함수 적용
    df['Times'] = df['Times'].apply(extract_time)

    
    

    def filter_tba(locations):
        # '/'로 위치 정보를 분리
        parts = locations.split('/')
        # 'TBA'가 아닌 위치만 필터링
        filtered_parts = [part for part in parts if part != 'TBA']

        unique_parts = list(set(filtered_parts))
        # 원래 순서대로 정렬
        unique_parts.sort(key=filtered_parts.index)
        # 필터링된 위치 정보를 '/'로 다시 연결
        return '/'.join(unique_parts)

# df['Buildings'] 열에 filter_tba 함수 적용
    df['Buildings'] = df['Buildings'].apply(filter_tba)
    df['Buildings'] = df['Buildings'].apply(lambda x: '' if 'TBA' in x else x)


    def format_days(days):
        # '/'를 기준으로 문자열을 분리
        parts = days.split('/')
        formatted_parts = []
        
        # 각 파트를 처리
        for part in parts:
            # 각 문자를 해당 요일의 약어로 변환
            part = part.replace('M', 'Mo').replace('W', 'We').replace('F', 'Fr').replace('Sat','Sa').replace('Sun','Su')
            # 쉼표로 각 약어를 구분
            formatted = ','.join([part[i:i+2] for i in range(0, len(part), 2)])
            formatted_parts.append(formatted)
        
        # '/'로 다시 연결
        return '/'.join(formatted_parts)

    df['Days'] = df['Days'].apply(format_days)


    def format_time_date(row):
        # Days와 Times를 '/'로 분리
        days_parts = row['Days'].split('/') if pd.notna(row['Days']) else []
        times_parts = row['Times'].split('/') if pd.notna(row['Times']) else []
        
        # 분할된 각 파트의 최소 길이를 확인하여 적절히 처리
        num_parts = min(len(days_parts), len(times_parts))
        
        # 각 파트별로 조합 생성
        formatted_parts = []
        for i in range(num_parts):
            if days_parts[i] != '' and times_parts[i] != '':
                formatted_part = f"{days_parts[i]}/{times_parts[i]}"
                if formatted_part not in formatted_parts:  # 중복 방지
                    formatted_parts.append(formatted_part)
        
        # 모든 조합을 쉼표로 구분하여 하나의 문자열로 반환
        return ', '.join(formatted_parts)

    # apply 함수를 사용하여 각 행에 format_time_date 함수 적용
    df['Time/Date'] = df.apply(format_time_date, axis=1)
    df['Time/Date'] = df['Time/Date'].apply(lambda x: '' if x == '/' else x)

    # DataFrame에 적용
    df.drop(columns=['Days','Times'], inplace=True)
    
    df.rename(columns={
        'Course Numbers': 'Section' ,
        'Course Names':'Course Name',
        'Credits': 'Credit',
        'Links to Professor Reviews':'Professor',
        'Buildings': 'Room',
        'CRNs':'Course Code'
    }, inplace=True)

    df = df[['Section','Course Name','Credit','Professor','Time/Date','Room','Course Code']] 
    
    # 출력 파일명 생성
    output_file_path = os.path.join(output_dir, os.path.basename(file_path))
    df.to_excel(output_file_path, index=False)
    print(f"완료된 학교: {output_file_path}")

# 파일이 있는 디렉토리
source_directory = '/Users/celine/Desktop/crw/Fall 2024_new'
output_directory = '/Users/celine/Desktop/crw/Fall 2024.final'

# 소스 디렉토리의 모든 파일을 순회
for filename in os.listdir(source_directory):
    file_path = os.path.join(source_directory, filename)
    if filename.endswith('.xlsx'):
        process_excel(file_path, output_directory)

