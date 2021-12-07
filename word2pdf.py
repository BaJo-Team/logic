import os
import win32com.client as win32
#word 단일 파일 변환
def Word_File_TO_PDF(input_file_path):
    input_file_name = input_file_path.split("\\")[-1]
    output_folder_path = os.path.join("C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트", input_file_name.split(".")[0] + ".pdf")  # 저장할 폴더 절대 경로

    file = open(output_folder_path, "w")
    file.close()

    application = win32.Dispatch('Word.Application')
    doc = application.Documents.Open(input_file_path)
    doc.SaveAs(output_folder_path, 17)
    doc.Close()
    application.Quit()
    return output_folder_path

#워드 다중파일 변환
def Word_MultiFile_To_PDF(input_folder_path):
    input_file_paths = os.listdir(input_folder_path)

    for input_file_name in input_file_paths:
        print(input_file_name)
        if not input_file_name.lower().endswith((".docx", ".doc")):
            continue

        if input_file_name.lower().startswith("~$"):
            continue

        output_folder_path = os.path.join("C:\\Users\\yeonsu\\Desktop\\단국대\\2021 1학년-2학기\\과제\\대학기초SW입문\\팀프로젝트",
                                          input_file_name.split(".")[0] + ".pdf")  # 저장을 할 폴더 절대 경로

        file = open(output_folder_path, "w")
        file.close()

        input_file_path = os.path.join(input_folder_path, input_file_name)

        application = win32.Dispatch('Word.Application')
        doc = application.Documents.Open(input_file_path)
        doc.SaveAs(output_folder_path, 17)
        doc.Close()
        application.Quit()
        output_file_paths = []
        for file in output_folder_path:
            output_file_paths = file
        return output_file_paths
