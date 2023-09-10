# MR_interview_schedule
카이스트 로봇 동아리 MR 면접 스케줄을 작성해주는 프로그램입니다. 만약, 수요일까지 면접을 본다면 주석 처리된 부분(wednedday 관련 부분)을 모두 활성화하고 사용하시면 됩니다.
알고리즘 작동 원리
각 시간대 별로 가능한 사람을 구한 후, 임의로 3명을 뽑아 시간표를 채워 나간다. 따라서, 최적의 경우라고 말할 수 없다.
랜덤으로 작동하기에, 작동할 때 마다 시간표 구성이 달라지며, 모든 사람이 시간표에 들어가지 않는 경우도 발생한다.
remain에 남은 사람이 표시되며, 시간표를 다시 한번 작동시키거나 사람에 의한 수정이 필요할 수도 있다.

엑셀 시트 예시와 코드를 함께 업로드 합니다.

# MR_interview_schedule
This is a program that creates an interview schedule for KAIST Robot Club MR. If you have an interview until Wednesday, you can activate and use all the commented parts (parts related to Wednesday).
How the algorithm works
After finding available people for each time zone, three people are randomly selected to fill out the timetable. Therefore, it cannot be said that this is the optimal case.
Because it operates randomly, the timetable composition changes each time it operates, and there are cases where not everyone fits into the timetable.
The remaining people are displayed in remain, and the timetable may need to be activated once again or modified by a human.

Upload the Excel sheet example and code together.
