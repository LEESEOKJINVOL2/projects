import tkinter as tk
from tkinter import filedialog
import csv
from itertools import product
import matplotlib.pyplot as plt
import pandas as pd

# 전역 변수로 df 선언
df = None

def choose_csv_file():
    global df
    
    # 파일 대화상자를 통해 CSV 파일 선택
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])

    if file_path:
        # CSV 파일 열기
        df = pd.read_csv(file_path)  # 데이터프레임에 CSV 파일 데이터 할당
        with open(file_path, 'r', newline='') as csv_file:
            csv_reader = csv.reader(csv_file)
            
            # CSV 파일의 첫 번째 행을 읽어 제목(title)을 가져옴
            header = next(csv_reader)
            
            # "TRASH", "LEVEL", "P/F", "NUMPACKET", "BLANK", "VALUE" 제목을 제외하고 데이터를 저장할 딕셔너리 초기화
            data_dict = {}
            selected_data = {}  # 선택한 데이터를 저장하는 딕셔너리
            
            for title in header:
                if title not in ["CH", "BAND","PAGAIN", "PWR_MODE", "TRASH", "LEVEL", "P/F", "NUMPACKET", "BLANK", "VALUE"]:
                    data_dict[title] = set()
                    selected_data[title] = set()  # 선택한 데이터를 저장하기 위해 초기화
            
            # 데이터를 수집
            for row in csv_reader:
                for title, value in zip(header, row):
                    if title in data_dict:
                        data_dict[title].add(value)
            
            # UI 생성
            root = tk.Tk()
            root.title("Auto Plot")
            
            # 체크박스 항목을 맨 위로 당기기 위한 LabelFrame 추가
            label_frame = tk.LabelFrame(root, text="Select Items", labelanchor="n")
            label_frame.grid(row=0, column=0, columnspan=len(data_dict), sticky="n")
            
            # 클릭 이벤트 핸들러
            def on_item_click(title, item):
                if item in selected_data[title]:
                    selected_data[title].remove(item)
                else:
                    selected_data[title].add(item)
            
            # 각 제목별로 체크박스를 생성하고 데이터를 보여주는 함수
            def show_data(title, label_frame):
                data = data_dict[title]
                for item in sorted(data):
                    var = tk.BooleanVar(value=True)  # 체크박스 초기 상태를 True로 설정
                    checkbox = tk.Checkbutton(label_frame, text=item, variable=var, command=lambda t=title, i=item: on_item_click(t, i))
                    checkbox.select()  # 체크박스를 선택 상태로 설정
                    checkbox.pack(anchor="w")
            
            # 각 제목별로 Label을 생성하여 제목을 표시하고 그 아래에 체크박스를 배치
            list_boxes = {}  # TITLE에 대한 Listbox를 저장하기 위한 딕셔너리
            for i, title in enumerate(data_dict):
                list_box = tk.LabelFrame(label_frame, text=title)
                list_box.pack(side="left", fill="y")
                list_boxes[title] = list_box  # LabelFrame에 대한 딕셔너리 키 설정
                show_data(title, list_box)  # 데이터를 초기에 출력
                    
            def create_and_save_graph(df, x_column, y_column, category):
                # 필터링: 선택한 카테고리를 기반으로 데이터 필터링
                filtered_data = df[(df['TEST_NAME'] == category[0]) & (df['PKT_TYPE'] == category[1])]
                print(filtered_data)
                if len(filtered_data) == 0:
                    print(f"No data for category: {category}")
                    return
                
                # 그래프 생성
                plt.figure()
                plt.scatter(filtered_data[x_column], filtered_data[y_column])
                plt.title(f"Category: {category[0]} @ {category[1]}")
                plt.xlabel(x_column)
                plt.ylabel(y_column)
                
                # 그래프를 이미지 파일로 저장
                output_file = f"C:/VS Code/Auto  Plot/png/category_{category[0]}_{category[1]}.png"
                plt.savefig(output_file)
                print(f"Saved graph as {output_file}")
                
            def plot_data():
                categories = []  # 카테고리 이름을 저장할 리스트
                selected_values = {}  # 각 제목별로 선택한 값들을 저장
                for title, values in data_dict.items():
                    selected_values[title] = selected_data[title]
                    if not selected_values[title]:
                        selected_values[title] = values
                for combo in product(*selected_values.values()):
                    category = "@".join(combo)
                    categories.append(category)
                # 각 카테고리 이름을 출력
                for i, category in enumerate(categories):
                    print(f"Category {i + 1}: {category}")    
                    # print(df.head())
                    create_and_save_graph(df, 'BAND', 'VALUE', category)
                    # 그래프 생성 및 저장 함수 호출

            plot_button = tk.Button(root, text="PLOT", command=plot_data)
            plot_button.grid(row=1, column=0)

            # QUIT 버튼을 추가하여 UI 종료
            quit_button = tk.Button(root, text="QUIT", command=root.quit)  # 프로그램을 종료하기 위해 root.quit 사용
            quit_button.grid(row=1, column=3, columnspan=len(data_dict))
            
            root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 루트 윈도우를 숨김
    choose_csv_file()