# hwp 변환
import os
import win32com.client as win32
import win32gui

# input_folder_path = "C:/Users/yeonsu/Desktop/단국대/2021 1학년-2학기/과제/대학기초SW입문/팀프로젝트"
# output_folder_path = "PDF 출력 폴더"
# input_file_paths = os.listdir(input_folder_path)    # 폴더 안에 있는 파일명 리스트로 출력
# print(input_file_paths)
#
# hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
# hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기
#
# print(hwnd)
#
# win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 숨겨줘. 0은 숨기기, 5는 보이기, 3은 풀스크린 등
# hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')  # 보안모듈 적용


# 엑셀 변환
# input_folder_path = "C:/Users/yeonsu/Desktop/단국대/2021 1학년-2학기/과제/대학기초SW입문/팀프로젝트"
# # output_folder_path = "PDF 출력 폴더"
# input_file_paths = os.listdir(input_folder_path)    # 폴더 안에 있는 파일명 -> list
# print(input_file_paths)
#
# for input_file_name in input_file_paths:
#
#     if not input_file_name.lower().endswith((".xlsx")):
#         continue
#
#     input_file_path = os.path.join(input_folder_path, input_file_name)
#
#     excel = win32.gencache.EnsureDispatch("Excel.Application")     # dispath를 쓰냐 / gencache.EnsureDispatch를 쓰냐
#     excel.Visible = True
#     xlsx = excel.Open(input_file_path)
#     print(xlsx)

# ppt 다중 파일 변환
# input_folder_path = "C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트"     # 파일이 있는 폴더 절대 경로
# output_folder_path = "C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트"    # 저장을 할 폴더 절대 경로
# input_file_paths = os.listdir(input_folder_path)
#
# for input_file_name in input_file_paths:
#
#     if not input_file_name.lower().endswith((".ppt", ".pptx")):
#         continue
#
#     input_file_path = os.path.join(input_folder_path, input_file_name)
#     application = win32.Dispatch("Powerpoint.Application")
#
#     presentation = application.Presentations.Open(input_file_path, ReadOnly=False)
#     presentation.SaveAs(os.path.join(output_folder_path, input_file_name.split('.')[0] + ".pdf"), 32)
#
#     application.Quit()

# ppt 단일 파일 변환

# output_folder_path = "C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트"    # 저장할 폴더 선택 기능 필요
#
# input_file_path = "C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트\\test.pptx"
# input_file_name = input_file_path.split("\\")[-1]
#
# application = win32.Dispatch("Powerpoint.Application")
#
# presentation = application.Presentations.Open(input_file_path, ReadOnly=False)
# presentation.SaveAs(os.path.join(output_folder_path, input_file_name.split(".")[0] + ".pdf"), 32)
#
# application.Quit()

# word 단일 파일 변환
# input_file_path = "C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트\\test3.docx"  # 파일 절대 경로
# input_file_name = input_file_path.split("\\")[-1]
# output_folder_path = os.path.join("C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트", input_file_name.split(".")[0] + ".pdf")   # 저장할 폴더 절대 경로
# print(output_folder_path)
# file = open(output_folder_path, "w")
# file.close()
#
# application = win32.Dispatch('Word.Application')
# doc = application.Documents.Open(input_file_path)
# doc.SaveAs(output_folder_path, 17)
# doc.Close()
# application.Quit()

#워드 다중파일 변환
input_folder_path = "C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트"  # 파일이 있는 폴더 절대 경로
input_file_paths = os.listdir(input_folder_path)

for input_file_name in input_file_paths:
    print(input_file_name)
    if not input_file_name.lower().endswith((".docx", ".doc")):
        continue

    if input_file_name.lower().startswith("~$"):
        continue

    output_folder_path = os.path.join("C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트", input_file_name.split(".")[0] + ".pdf")  # 저장을 할 폴더 절대 경로

    file = open(output_folder_path, "w")
    file.close()

    input_file_path = os.path.join(input_folder_path, input_file_name)

    application = win32.Dispatch('Word.Application')
    doc = application.Documents.Open(input_file_path)
    doc.SaveAs(output_folder_path, 17)
    doc.Close()
    application.Quit()