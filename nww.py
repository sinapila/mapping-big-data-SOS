from difflib import SequenceMatcher
import pandas as pd
import numpy as np
import arabic_reshaper
from bidi.algorithm import get_display

file = 'sina222.xlsx'


par = False


w_count = 0

def sm(s, t, ratio_calc = True):

    # Initialize matrix of zeros
    rows = len(s)+1
    cols = len(t)+1
    distance = np.zeros((rows,cols),dtype = int)

    # Populate matrix of zeros with the indeces of each character of both strings
    for i in range(1, rows):
        for k in range(1,cols):
            distance[i][0] = i
            distance[0][k] = k

    # Iterate over the matrix to compute the cost of deletions,insertions and/or substitutions    
    for col in range(1, cols):
        for row in range(1, rows):
            if s[row-1] == t[col-1]:
                cost = 0 # If the characters are the same in the two strings in a given position [i,j] then the cost is 0
            else:
                # In order to align the results with those of the Python Levenshtein package, if we choose to calculate the ratio
                # the cost of a substitution is 2. If we calculate just distance, then the cost of a substitution is 1.
                if ratio_calc == True:
                    cost = 2
                else:
                    cost = 1
            distance[row][col] = min(distance[row-1][col] + 1,      # Cost of deletions
                                 distance[row][col-1] + 1,          # Cost of insertions
                                 distance[row-1][col-1] + cost)     # Cost of substitutions
    if ratio_calc == True:
        # Computation of the Levenshtein Distance Ratio
        Ratio = ((len(s)+len(t)) - distance[row][col]) / (len(s)+len(t))
        return Ratio
    else:
        # print(distance) # Uncomment if you want to see the matrix showing how the algorithm computes the cost of deletions,
        # insertions and/or substitutions
        # This is the minimum number of edits needed to convert string a to string b
        return "The strings are {} edits away".format(distance[row][col])
    
def load_data(file):
    global vazarat,coamk

    xl = pd.ExcelFile(file)

    vazarat = xl.parse('وزارت')
    coamk = xl.parse('کمک رسان')
    print("done !")

def make_siam_colum():
    fake_data = []
    for i in range(578):
        fake_data.append("")

    coamk["شناسه سیام"] = fake_data

def preprocessing():

    coamk["استان"] = coamk["استان"].apply(lambda x : str(x))
    coamk["نام مرکز"] = coamk["نام مرکز"].apply(lambda x : str(x))
    coamk["شهر"] = coamk["شهر"].apply(lambda x : str(x))
    coamk["شناسه سیام"] = coamk["شناسه سیام"].apply(lambda x : str(x))

    vazarat["استان"] = vazarat["استان"].apply(lambda x : x.replace("آذربایجان","").replace("دندانپزشکی","").replace("دانشکده","").replace("جراحی","").replace("دکتر","")
                                              .replace("حضرت","").replace("مرکز","").replace("-","").replace("فیزیوتراپی","").replace(")","").replace("(","")
                                              .replace("مطب","").replace("_","").replace("/","").replace("شبانه روزی","").replace("طبی","").replace("رادیولوژی","")
                                              .replace("آذربایجان شرقی","").replace("درمانگاه","").replace("بیمارستان","").replace("داروخانه","").replace("عینک سازی","")
                                              .replace("آزمایشگاه","").replace("پزشکی","").replace("اذربایجان","").replace("خراسان ","").replace("ي","ی")
                                              .replace("مرکزﻱ",""))
    coamk["استان"] = coamk["استان"].apply(lambda x : x.replace("آذربایجان","").replace("دندانپزشکی","").replace("دانشکده","").replace("جراحی","").replace("دکتر","")
                                          .replace("حضرت","").replace("مرکز","").replace("-","").replace("فیزیوتراپی","").replace(")","").replace("(","").replace("مطب","")
                                          .replace("_","").replace("/","").replace("شبانه روزی","").replace("طبی","").replace("رادیولوژی","").replace("آذربایجان شرقی","")
                                          .replace("درمانگاه","").replace("بیمارستان","").replace("داروخانه","").replace("عینک سازی","").replace("آزمایشگاه","").replace("پزشکی","")
                                          .replace("اذربایجان","").replace("خراسان ","").replace("اذربايجان","").replace("استان","").replace("ي","ی"))
    coamk["نام مرکز"] = coamk["نام مرکز"].apply(lambda x : x.replace("دکتر","").replace("مراغه","").replace("تصویر برداری","").replace("دندانپزشکی","").replace("دانشکده","")
                                                .replace("جراحی","").replace("دکتر","").replace("حضرت","").replace("مرکز","").replace("-","").replace("فیزیوتراپی","")
                                                .replace(")","").replace("(","").replace("مطب","").replace("_","").replace("/","").replace("شبانه روزی","")
                                                .replace("طبی","").replace("رادیولوژی","").replace("آذربایجان شرقی","").replace("درمانگاه","").replace("بیمارستان","")
                                                .replace("داروخانه","").replace("عینک سازی","").replace("آزمایشگاه","").replace("پزشکی","").replace("اذربایجان","")
                                                .replace("-"," ").replace(")"," ").replace("("," ").replace("آذربایجان","").replace("خراسان","").replace("ي","ی")
                                                .replace("تبريز","").replace("عجبشیر","").replace("اذربايجان شرقي","").replace("مازندران","").replace("همدان","")
                                                .replace("گیلان","").replace("کرمانشاه","").replace("ﺮﯿﺸﺒﺠﻋ","").replace("مرند","").replace("تصویربرداری","")
                                                .replace("تخت خوابی","").replace("تبریز","").replace("بستری","").replace("اردبیل","").replace("آزمایشگاه","")
                                                .replace("دانشکده","").replace("شرقی","").replace("شرقي","").replace("غربی","").replace("بستان آباد","").replace("ارومیه","")
                                                .replace("اصفهان","").replace("لنجان","").replace("اصفهان","").replace("تیران و کرون","").replace("اصفهان","").replace("فلاورجان","")
                                                .replace("ايذه","").replace("باغ  ملک","").replace("انديمشک","").replace("آران و بیدگل","").replace("خمینی شهر","").replace("مبارکه","")
                                                .replace("شوش","").replace("شوشتر","").replace("بندرماهشهر","").replace("اميديه","").replace("بندرامام  خميني","").replace("بهبهان","")
                                                .replace("آبادان","").replace("دزفول","").replace("خرمشهر","").replace("کردستان","").replace("اغاجاري","").replace("مشهد","").replace("میاندوآب","")
                                                .replace("مرکزي","").replace("میانه","").replace(",","").replace("شبستر","").replace("سونوگرافی","").replace("پاتولوژی","")
                                                .replace("پاراکلینیکی","").replace("تخصصی",""))
    

    coamk["شهر"] = coamk["شهر"].apply(lambda x : x.replace("شهر ري","ری").replace("مرکزي  اردبيل","اردبیل").replace("مرکزي","").replace("مرکزی  اردبيل","اردبیل").replace("مرکزی",""))
    

    vazarat["نام"] = vazarat["نام"].apply(lambda x : x.replace("دکتر",""))

    for i in coamk.iloc:
        i["نام مرکز"] = i["نام مرکز"].replace(i["استان"],"").replace(i["شهر"],"")
    


    print(coamk["استان"] )
    
def initial_final_data():
    global final_data

    data = {'کد مرکز': [],
	'نام مرکز': [],
    'نوع': [],
	'استان': [],
	'شهر': [],
    'شناسه سیام': [],}

    final_data = pd.DataFrame(data)

def convert(text):
    reshaped_text = arabic_reshaper.reshape(text)
    converted = get_display(reshaped_text)
    return converted



# def write_on_df

load_data(file)

make_siam_colum()

preprocessing()

initial_final_data()
# print(coamk.iloc[2])


print(vazarat)
print(coamk)


max_similarity = 0
max_similarity_row_vazarat_item = []
max_similarity_row_coamk_item = []


last = vazarat.iloc[0]["استان"]
last_2 = coamk.iloc[0]["استان"]
# print(vazarat.iloc[3]["id"])
count = 0
countt = 0
for coamk_item in coamk.iloc:
    
    max_similarity = 0
    max_similarity_row_vazarat_item = []
    max_similarity_row_coamk_item = []
    countt += 1
    print(f"{countt} / 7523")


    try:



        for vazarat_item in vazarat.iloc:


# and (sm(str(coamk_item["شهر"]),str(vazarat_item["شهرستان"]))>0.5)
            
            if (sm(coamk_item["استان"],vazarat_item["استان"].replace("استان","")) > 0.5)  and (sm(str(coamk_item["نوع"]),str(vazarat_item["نوع"]))>0.66):
                # print(sm(coamk_item["استان"],vazarat_item["استان"]),coamk_item["استان"],vazarat_item["استان"])

                # print(convert(str(coamk_item['نام مرکز'])),convert(str(coamk_item['نام مرکز']).replace( str(vazarat_item['شهرستان']),"")),convert(str(vazarat_item['نام'])),convert(str(vazarat_item['نام'])))

                similarity = sm(str(coamk_item['نام مرکز']).replace( str(vazarat_item['شهرستان']),""),str(vazarat_item['نام']))


                if similarity > max_similarity :



                    max_similarity = similarity
                    max_similarity_row_vazarat_item = vazarat_item
                    max_similarity_row_coamk_item = coamk_item
                    # print(max_similarity)

                # print(similarity)

            if last != vazarat_item["استان"]:
                # last = vazarat_item["استان"]
                break

            last = vazarat_item["استان"]


        
        if max_similarity > 0.66:

            print(convert(str(max_similarity_row_coamk_item['نام مرکز']) + " -----> " + str(max_similarity_row_vazarat_item['نام']) + "با  " + str(max_similarity)+" Ok "),convert(max_similarity_row_coamk_item["شهر"]),convert(max_similarity_row_vazarat_item["شهرستان"]))
            max_similarity_row_coamk_item["شناسه سیام"] = max_similarity_row_vazarat_item["شناسه سیام"] 

            final_data = final_data.append(max_similarity_row_coamk_item, ignore_index=True)
            
        

        else:
            try:
                max_similarity_row_coamk_item["شناسه سیام"] = ""
                print(convert(str(max_similarity_row_coamk_item['نام مرکز']) + " -----> " + str(max_similarity_row_vazarat_item['نام']) + " با  " + str(max_similarity)+" ERR "),convert(max_similarity_row_coamk_item["شهر"]),convert(max_similarity_row_vazarat_item["شهرستان"]))

            except:
                pass
            final_data = final_data.append(max_similarity_row_coamk_item, ignore_index=True)



    except:
        print("Big ERR")
        final_data = final_data.append({'کد مرکز': "",
        'نام مرکز': "",
        'نوع': "",
        'استان': "",
        'شهر': "",
        'شناسه سیام': "",}, ignore_index=True)



    if last_2 != coamk_item["استان"]:
        vazarat.drop(vazarat.loc[vazarat["استان"] == last].index, inplace=True)
        # writer = pd.ExcelWriter(str(last_2)+'.xlsx', engine='xlsxwriter')
        # final_data.to_excel(writer, 'کمک رسان')
        # writer.save()
        # data = {'کد مرکز': [],
        # 'نام مرکز': [],
        # 'نوع': [],
        # 'استان': [],
        # 'شهر': [],
        # 'شناسه سیام': [],}
        # final_data = pd.DataFrame(data)
        try:
            last = vazarat.iloc[0]["استان"]
        except:
            pass

    last_2 = coamk_item["استان"]


    
writer = pd.ExcelWriter("finisher"+'.xlsx', engine='xlsxwriter')
final_data.to_excel(writer, 'کمک رسان')
writer.save()


    