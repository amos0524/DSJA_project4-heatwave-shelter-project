// 출처1: data/heat_patients_final/온열질환_사망자_시군구별_통합_2021-2026.csv (2021~2025년 구간만, 시군구 사망자수)
// 출처2: project4_material/project4_chapter2/project4_(2025추가본)..._20260808.xlsx Sheet1 (시도별 농작업 현장 온열질환자수 = 실외_실외작업장+실외_논밭+실외_산+실내_비닐하우스 합)
// 주의1: 이 파일의 '2025' 탭은 Sheet1의 2025년 행과 값이 완전히 동일한 중복 탭이라 사용하지 않았다(둘 다 쓰면 값이 두 번 잡힘).
// 주의2: 사망자수는 장소(농작업 여부) 구분이 없는 시군구 전체 사망자다(원자료에 사망자 장소 컬럼 없음). 농작업 여부는 환자수(farm_patients) 기준으로만 판단 가능.
var REGION5 = {
 "sgg_deaths_2021_2025": [
  {
   "year": 2021,
   "sido": "강원특별자치도",
   "sgg": "강릉시",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "강원특별자치도",
   "sgg": "고성군",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "강원특별자치도",
   "sgg": "원주시",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "강원특별자치도",
   "sgg": "정선군",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "경기도",
   "sgg": "양주시",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "경기도",
   "sgg": "의정부시",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "경상남도",
   "sgg": "남해군",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "경상북도",
   "sgg": "구미시",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "경상북도",
   "sgg": "봉화군",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "부산광역시",
   "sgg": "해운대구",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "서울특별시",
   "sgg": "양천구",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "서울특별시",
   "sgg": "종로구",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "서울특별시",
   "sgg": "중랑구",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "인천광역시",
   "sgg": "남동구",
   "deaths": 2
  },
  {
   "year": 2021,
   "sido": "인천광역시",
   "sgg": "서해구",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "전남광주통합특별시",
   "sgg": "북구",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "충청남도",
   "sgg": "공주시",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "충청남도",
   "sgg": "부여군",
   "deaths": 1
  },
  {
   "year": 2021,
   "sido": "충청북도",
   "sgg": "단양군",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "경기도",
   "sgg": "부천시",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "경기도",
   "sgg": "파주시",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "경상남도",
   "sgg": "창녕군",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "경상남도",
   "sgg": "함양군",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "인천광역시",
   "sgg": "강화군",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "인천광역시",
   "sgg": "부평구",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "전라북도",
   "sgg": "김제시",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "충청남도",
   "sgg": "논산시",
   "deaths": 1
  },
  {
   "year": 2022,
   "sido": "충청북도",
   "sgg": "청주시 청원구",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "경상남도",
   "sgg": "거제시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "경상남도",
   "sgg": "남해군",
   "deaths": 2
  },
  {
   "year": 2023,
   "sido": "경상남도",
   "sgg": "밀양시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "경상남도",
   "sgg": "창녕군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "경상남도",
   "sgg": "함양군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "경상북도",
   "sgg": "문경시",
   "deaths": 2
  },
  {
   "year": 2023,
   "sido": "경상북도",
   "sgg": "상주시",
   "deaths": 2
  },
  {
   "year": 2023,
   "sido": "대구광역시",
   "sgg": "북구",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "부산광역시",
   "sgg": "서구",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "서울특별시",
   "sgg": "송파구",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "울산광역시",
   "sgg": "남구",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전남광주통합특별시",
   "sgg": "순천시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전남광주통합특별시",
   "sgg": "여수시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전남광주통합특별시",
   "sgg": "영암군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전라북도",
   "sgg": "군산시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전라북도",
   "sgg": "무주군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전라북도",
   "sgg": "정읍시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "전라북도",
   "sgg": "진안군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "충청남도",
   "sgg": "당진시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "충청남도",
   "sgg": "부여군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "충청남도",
   "sgg": "서천군",
   "deaths": 3
  },
  {
   "year": 2023,
   "sido": "충청남도",
   "sgg": "아산시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "충청남도",
   "sgg": "예산군",
   "deaths": 2
  },
  {
   "year": 2023,
   "sido": "충청북도",
   "sgg": "보은군",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "충청북도",
   "sgg": "제천시",
   "deaths": 1
  },
  {
   "year": 2023,
   "sido": "충청북도",
   "sgg": "충주시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "강원특별자치도",
   "sgg": "삼척시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "강원특별자치도",
   "sgg": "인제군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경기도",
   "sgg": "부천시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경기도",
   "sgg": "여주시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경상남도",
   "sgg": "밀양시",
   "deaths": 2
  },
  {
   "year": 2024,
   "sido": "경상남도",
   "sgg": "창녕군",
   "deaths": 2
  },
  {
   "year": 2024,
   "sido": "경상남도",
   "sgg": "창원시 마산합포구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경상남도",
   "sgg": "함안군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경상북도",
   "sgg": "경주시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경상북도",
   "sgg": "상주시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "경상북도",
   "sgg": "포항시 북구",
   "deaths": 3
  },
  {
   "year": 2024,
   "sido": "대구광역시",
   "sgg": "군위군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "부산광역시",
   "sgg": "연제구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "서울특별시",
   "sgg": "도봉구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "서울특별시",
   "sgg": "중랑구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "인천광역시",
   "sgg": "미추홀구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전남광주통합특별시",
   "sgg": "고흥군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전남광주통합특별시",
   "sgg": "서구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전남광주통합특별시",
   "sgg": "신안군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전남광주통합특별시",
   "sgg": "장성군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전남광주통합특별시",
   "sgg": "장흥군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전남광주통합특별시",
   "sgg": "해남군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "전라북도",
   "sgg": "진안군",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "제주특별자치도",
   "sgg": "서귀포시",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "충청남도",
   "sgg": "예산군",
   "deaths": 2
  },
  {
   "year": 2024,
   "sido": "충청남도",
   "sgg": "천안시 서북구",
   "deaths": 1
  },
  {
   "year": 2024,
   "sido": "충청남도",
   "sgg": "태안군",
   "deaths": 2
  },
  {
   "year": 2024,
   "sido": "충청북도",
   "sgg": "제천시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경기도",
   "sgg": "성남시 분당구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경기도",
   "sgg": "이천시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경기도",
   "sgg": "파주시",
   "deaths": 2
  },
  {
   "year": 2025,
   "sido": "경기도",
   "sgg": "포천시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경기도",
   "sgg": "화성시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경상남도",
   "sgg": "진주시",
   "deaths": 2
  },
  {
   "year": 2025,
   "sido": "경상남도",
   "sgg": "통영시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경상북도",
   "sgg": "경산시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경상북도",
   "sgg": "봉화군",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경상북도",
   "sgg": "영덕군",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "경상북도",
   "sgg": "영천시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "부산광역시",
   "sgg": "부산진구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "서울특별시",
   "sgg": "강동구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "서울특별시",
   "sgg": "관악구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "서울특별시",
   "sgg": "중랑구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "울산광역시",
   "sgg": "남구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "인천광역시",
   "sgg": "부평구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "인천광역시",
   "sgg": "서해구",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "전남광주통합특별시",
   "sgg": "목포시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "전남광주통합특별시",
   "sgg": "보성군",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "전남광주통합특별시",
   "sgg": "신안군",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "전라북도",
   "sgg": "진안군",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "충청남도",
   "sgg": "공주시",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "충청남도",
   "sgg": "당진시",
   "deaths": 2
  },
  {
   "year": 2025,
   "sido": "충청남도",
   "sgg": "부여군",
   "deaths": 1
  },
  {
   "year": 2025,
   "sido": "충청남도",
   "sgg": "청양군",
   "deaths": 1
  }
 ],
 "sido_farm_patients_2021_2025": [
  {
   "year": 2021,
   "sido": "강원",
   "farm_patients": 33,
   "total_patients": 55
  },
  {
   "year": 2021,
   "sido": "경기",
   "farm_patients": 145,
   "total_patients": 271
  },
  {
   "year": 2021,
   "sido": "경남",
   "farm_patients": 71,
   "total_patients": 126
  },
  {
   "year": 2021,
   "sido": "경북",
   "farm_patients": 72,
   "total_patients": 124
  },
  {
   "year": 2021,
   "sido": "광주",
   "farm_patients": 18,
   "total_patients": 32
  },
  {
   "year": 2021,
   "sido": "대구",
   "farm_patients": 18,
   "total_patients": 32
  },
  {
   "year": 2021,
   "sido": "대전",
   "farm_patients": 17,
   "total_patients": 27
  },
  {
   "year": 2021,
   "sido": "부산",
   "farm_patients": 20,
   "total_patients": 45
  },
  {
   "year": 2021,
   "sido": "서울",
   "farm_patients": 34,
   "total_patients": 121
  },
  {
   "year": 2021,
   "sido": "세종",
   "farm_patients": 11,
   "total_patients": 14
  },
  {
   "year": 2021,
   "sido": "울산",
   "farm_patients": 23,
   "total_patients": 32
  },
  {
   "year": 2021,
   "sido": "인천",
   "farm_patients": 44,
   "total_patients": 92
  },
  {
   "year": 2021,
   "sido": "전남",
   "farm_patients": 79,
   "total_patients": 110
  },
  {
   "year": 2021,
   "sido": "전북",
   "farm_patients": 54,
   "total_patients": 93
  },
  {
   "year": 2021,
   "sido": "제주",
   "farm_patients": 45,
   "total_patients": 65
  },
  {
   "year": 2021,
   "sido": "충남",
   "farm_patients": 50,
   "total_patients": 83
  },
  {
   "year": 2021,
   "sido": "충북",
   "farm_patients": 36,
   "total_patients": 54
  },
  {
   "year": 2022,
   "sido": "강원",
   "farm_patients": 37,
   "total_patients": 62
  },
  {
   "year": 2022,
   "sido": "경기",
   "farm_patients": 199,
   "total_patients": 353
  },
  {
   "year": 2022,
   "sido": "경남",
   "farm_patients": 86,
   "total_patients": 152
  },
  {
   "year": 2022,
   "sido": "경북",
   "farm_patients": 76,
   "total_patients": 142
  },
  {
   "year": 2022,
   "sido": "광주",
   "farm_patients": 11,
   "total_patients": 20
  },
  {
   "year": 2022,
   "sido": "대구",
   "farm_patients": 13,
   "total_patients": 29
  },
  {
   "year": 2022,
   "sido": "대전",
   "farm_patients": 10,
   "total_patients": 24
  },
  {
   "year": 2022,
   "sido": "부산",
   "farm_patients": 31,
   "total_patients": 53
  },
  {
   "year": 2022,
   "sido": "서울",
   "farm_patients": 30,
   "total_patients": 110
  },
  {
   "year": 2022,
   "sido": "세종",
   "farm_patients": 15,
   "total_patients": 26
  },
  {
   "year": 2022,
   "sido": "울산",
   "farm_patients": 9,
   "total_patients": 26
  },
  {
   "year": 2022,
   "sido": "인천",
   "farm_patients": 34,
   "total_patients": 69
  },
  {
   "year": 2022,
   "sido": "전남",
   "farm_patients": 79,
   "total_patients": 124
  },
  {
   "year": 2022,
   "sido": "전북",
   "farm_patients": 67,
   "total_patients": 97
  },
  {
   "year": 2022,
   "sido": "제주",
   "farm_patients": 59,
   "total_patients": 93
  },
  {
   "year": 2022,
   "sido": "충남",
   "farm_patients": 88,
   "total_patients": 135
  },
  {
   "year": 2022,
   "sido": "충북",
   "farm_patients": 30,
   "total_patients": 49
  },
  {
   "year": 2023,
   "sido": "강원",
   "farm_patients": 56,
   "total_patients": 104
  },
  {
   "year": 2023,
   "sido": "경기",
   "farm_patients": 330,
   "total_patients": 683
  },
  {
   "year": 2023,
   "sido": "경남",
   "farm_patients": 128,
   "total_patients": 226
  },
  {
   "year": 2023,
   "sido": "경북",
   "farm_patients": 140,
   "total_patients": 255
  },
  {
   "year": 2023,
   "sido": "광주",
   "farm_patients": 33,
   "total_patients": 64
  },
  {
   "year": 2023,
   "sido": "대구",
   "farm_patients": 22,
   "total_patients": 59
  },
  {
   "year": 2023,
   "sido": "대전",
   "farm_patients": 21,
   "total_patients": 41
  },
  {
   "year": 2023,
   "sido": "부산",
   "farm_patients": 38,
   "total_patients": 94
  },
  {
   "year": 2023,
   "sido": "서울",
   "farm_patients": 48,
   "total_patients": 201
  },
  {
   "year": 2023,
   "sido": "세종",
   "farm_patients": 18,
   "total_patients": 27
  },
  {
   "year": 2023,
   "sido": "울산",
   "farm_patients": 39,
   "total_patients": 68
  },
  {
   "year": 2023,
   "sido": "인천",
   "farm_patients": 48,
   "total_patients": 113
  },
  {
   "year": 2023,
   "sido": "전남",
   "farm_patients": 143,
   "total_patients": 222
  },
  {
   "year": 2023,
   "sido": "전북",
   "farm_patients": 88,
   "total_patients": 207
  },
  {
   "year": 2023,
   "sido": "제주",
   "farm_patients": 69,
   "total_patients": 98
  },
  {
   "year": 2023,
   "sido": "충남",
   "farm_patients": 115,
   "total_patients": 205
  },
  {
   "year": 2023,
   "sido": "충북",
   "farm_patients": 92,
   "total_patients": 151
  },
  {
   "year": 2024,
   "sido": "강원",
   "farm_patients": 79,
   "total_patients": 160
  },
  {
   "year": 2024,
   "sido": "경기",
   "farm_patients": 367,
   "total_patients": 767
  },
  {
   "year": 2024,
   "sido": "경남",
   "farm_patients": 216,
   "total_patients": 377
  },
  {
   "year": 2024,
   "sido": "경북",
   "farm_patients": 148,
   "total_patients": 290
  },
  {
   "year": 2024,
   "sido": "광주",
   "farm_patients": 42,
   "total_patients": 70
  },
  {
   "year": 2024,
   "sido": "대구",
   "farm_patients": 19,
   "total_patients": 67
  },
  {
   "year": 2024,
   "sido": "대전",
   "farm_patients": 25,
   "total_patients": 69
  },
  {
   "year": 2024,
   "sido": "부산",
   "farm_patients": 54,
   "total_patients": 129
  },
  {
   "year": 2024,
   "sido": "서울",
   "farm_patients": 82,
   "total_patients": 235
  },
  {
   "year": 2024,
   "sido": "세종",
   "farm_patients": 11,
   "total_patients": 29
  },
  {
   "year": 2024,
   "sido": "울산",
   "farm_patients": 65,
   "total_patients": 109
  },
  {
   "year": 2024,
   "sido": "인천",
   "farm_patients": 89,
   "total_patients": 205
  },
  {
   "year": 2024,
   "sido": "전남",
   "farm_patients": 221,
   "total_patients": 407
  },
  {
   "year": 2024,
   "sido": "전북",
   "farm_patients": 116,
   "total_patients": 227
  },
  {
   "year": 2024,
   "sido": "제주",
   "farm_patients": 74,
   "total_patients": 123
  },
  {
   "year": 2024,
   "sido": "충남",
   "farm_patients": 143,
   "total_patients": 244
  },
  {
   "year": 2024,
   "sido": "충북",
   "farm_patients": 122,
   "total_patients": 196
  },
  {
   "year": 2025,
   "sido": "강원",
   "farm_patients": 72,
   "total_patients": 172
  },
  {
   "year": 2025,
   "sido": "경기",
   "farm_patients": 469,
   "total_patients": 978
  },
  {
   "year": 2025,
   "sido": "경남",
   "farm_patients": 192,
   "total_patients": 382
  },
  {
   "year": 2025,
   "sido": "경북",
   "farm_patients": 229,
   "total_patients": 436
  },
  {
   "year": 2025,
   "sido": "광주",
   "farm_patients": 36,
   "total_patients": 72
  },
  {
   "year": 2025,
   "sido": "대구",
   "farm_patients": 70,
   "total_patients": 140
  },
  {
   "year": 2025,
   "sido": "대전",
   "farm_patients": 22,
   "total_patients": 52
  },
  {
   "year": 2025,
   "sido": "부산",
   "farm_patients": 84,
   "total_patients": 171
  },
  {
   "year": 2025,
   "sido": "서울",
   "farm_patients": 74,
   "total_patients": 378
  },
  {
   "year": 2025,
   "sido": "세종",
   "farm_patients": 17,
   "total_patients": 33
  },
  {
   "year": 2025,
   "sido": "울산",
   "farm_patients": 98,
   "total_patients": 188
  },
  {
   "year": 2025,
   "sido": "인천",
   "farm_patients": 85,
   "total_patients": 280
  },
  {
   "year": 2025,
   "sido": "전남",
   "farm_patients": 255,
   "total_patients": 381
  },
  {
   "year": 2025,
   "sido": "전북",
   "farm_patients": 127,
   "total_patients": 248
  },
  {
   "year": 2025,
   "sido": "제주",
   "farm_patients": 69,
   "total_patients": 107
  },
  {
   "year": 2025,
   "sido": "충남",
   "farm_patients": 130,
   "total_patients": 227
  },
  {
   "year": 2025,
   "sido": "충북",
   "farm_patients": 116,
   "total_patients": 215
  }
 ]
};
