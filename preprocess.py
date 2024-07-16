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
    
    df['Credits'] = df['Credits'].apply(lambda x: '' if x == 'TBA' else x)
    df['Links to Professor Reviews'] = df['Links to Professor Reviews'].apply(lambda x: '' if x == '- -' else x)
    def filter_tba_time_segments(time_string):
    # '/'로 구분된 시간 구간을 분리
        segments = time_string.split('/')
        # 'TBA-TBA'를 제외한 시간 구간만 필터링
        filtered_segments = [segment for segment in segments if segment != 'TBA-TBA']
        # 결과를 '/'로 다시 연결
        return '/'.join(filtered_segments)

    # DataFrame의 'Times' 칼럼에 함수 적용 예시
    df['Times'] = df['Times'].apply(filter_tba_time_segments)

    
    

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
            part = part.replace('M', 'Mo').replace('W', 'We').replace(' thru ','Th').replace('F', 'Fr').replace('Sat','Sa').replace('Sun','Su')
            # 쉼표로 각 약어를 구분
            formatted = ','.join([part[i:i+2] for i in range(0, len(part), 2)])
            formatted_parts.append(formatted)
            

        
        # '/'로 다시 연결
        return '/'.join(formatted_parts)

    df['Days'] = df['Days'].apply(format_days)


    
    def convert_to_24h(time_str):
        # 시간과 분, 그리고 AM/PM을 추출
        time_str.split('-')
        match = re.match(r'(\d+):(\d+)(am|pm)', time_str)
        if not match:
            return time_str  # 일치하지 않으면 원래 문자열을 반환

        hour, minute, period = match.groups()
        hour = int(hour)
        minute = int(minute)

        # PM인 경우 12를 더하고, 12:00 PM은 예외적으로 12로 처리
        if period == 'pm':
            if hour != 12:
                hour += 12
        elif period == 'am' and hour == 12:
            hour = 0  # 12 AM은 0시

        return f"{hour:02}:{minute:02}"

    def convert_time_range(time_range):
        if '-' in time_range:
            start_time, end_time = time_range.split('-')
            start_24h = convert_to_24h(start_time.strip())
            end_24h = convert_to_24h(end_time.strip())
        else:
            # '-'가 없는 경우, 동일한 시간으로 시작과 끝을 처리
            start_24h = convert_to_24h(time_range.strip())
            end_24h = start_24h  # 끝 시간을 시작 시간과 동일하게 설정

        return f"{start_24h}-{end_24h}"
    print(df['Days'].iloc[6055:6066])

    def formatted_parts(row):
        days_list = row['Days'].split('/')
        days_list = row['Days'].split(',')
        time_ranges = row['Times'].split('/')
        results = []


        # 모든 요일에 대해 각각의 시간 매핑
        for day in days_list:
            for time_range in time_ranges:
                if '-' in time_range:
                    start_time, end_time = time_range.split('-')
                    start_24h = convert_to_24h(start_time.strip())
                    end_24h = convert_to_24h(end_time.strip())
                    results.append(f"{day}/{start_24h}-{end_24h}")

        # 중복 제거
        unique_results = sorted(set(results))
        return ', '.join(unique_results)


    df['Time/Date'] = df.apply(formatted_parts, axis=1)

    


  


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
    new_filename = os.path.splitext(os.path.basename(file_path))[0].split('_')[0] + "_Fall2024.xlsx"
    output_file_path = os.path.join(output_dir, new_filename)
    df.to_excel(output_file_path, index=False)
    print(f"완료된 학교: {output_file_path}")

# 파일이 있는 디렉토리
source_directory = '/Users/celine/Desktop/crw/Fall 2024.final/ready'
output_directory = '/Users/celine/Desktop/crw'

# 소스 디렉토리의 모든 파일을 순회
for filename in os.listdir(source_directory):
    file_path = os.path.join(source_directory, filename)
    if filename.endswith('.xlsx'):
        process_excel(file_path, output_directory)

