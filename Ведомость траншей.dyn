# Load the Python Standard and DesignScript Libraries
import sys
import clr
import re
import math

# Add Assemblies for AutoCAD and Civil3D
clr.AddReference('AcMgd')
clr.AddReference('AcCoreMgd')
clr.AddReference('AcDbMgd')
clr.AddReference('AecBaseMgd')
clr.AddReference('AecPropDataMgd')
clr.AddReference('AeccDbMgd')

# Import references from AutoCAD
from Autodesk.AutoCAD.Runtime import *
from Autodesk.AutoCAD.ApplicationServices import *
from Autodesk.AutoCAD.EditorInput import *
from Autodesk.AutoCAD.DatabaseServices import *
from Autodesk.AutoCAD.Geometry import *

# Import references from Civil3D
from Autodesk.Civil.ApplicationServices import *
from Autodesk.Civil.DatabaseServices import *

# обязательно нужно для работы с таблицами
from Autodesk.AutoCAD.DatabaseServices import Table as AcadTable

# !!! Это обязательно для обращения к наборам характеристик !!!
# Import references for PropertySets 
#from Autodesk.Aec.PropertyData import *
from Autodesk.Aec.PropertyData.DatabaseServices import *

# The inputs to this node will be stored as a list in the IN variables.
dataEnteringNode = IN

adoc = Application.DocumentManager.MdiActiveDocument
editor = adoc.Editor
civdoc = CivilApplication.ActiveDocument


TRENCH_STYLE_ES = "ЭС" # стиль траншеи ЭС
TRENCH_STYLE_EN = "ЭН" # стиль траншеи ЭН

TRENCH_TYPE = "Тип траншеи" # тип траншеи
TRENCH_HEIGHT = "Глубина Н м" # глубина траншеи
TRENCH_WIDTH = "Ширина В м" # ширина траншеи
TRENCH_SAND_BACKFILL = 'Песчаная засыпка h м' # высота песчанной засыпки
TRENCH_SAND_CUSHION = 'Песчаная подушка h м' # высота песчанной подушки
TRENCH_SLOPE_STEEPNESS = 'Угол крутизны откоса' # крутизна откоса

TRENCH_STYLE = "102_ось траншеи" # для поиска трасс с стилем трашеи
PATTERN_NAME_TRENCH = r"\_ТК\-(?:\d+?\.\d+|\d+\Z)" # шаблон имени трассы траншеи
########

PATTERN_TYPE_TRENCH = r"ТК\-\d+" # шаблон имени типа траншеи
SET_CHARACTERISTICS = 'ASML_Траншеи' # имя набора характеристик, с которым мы работаем
KF_SAND_COMPACTION = 1.1 # кф уплотнения песка

##################
# данные для наименования столбцов в таблице
TYPE_TR = "Тип"
LENGTH_ALIG = "Длина трассы"
SAND_CUSHION = "Постель"
SAND_BACKFILL = "Песок засыпки"
SOIL_BACKFILL = "Засыпка грунтом"
################

def create_table(volume_sand_cushion, name_trench_style):
#Данные для зполнения ячеек
    data = volume_sand_cushion
    
    # volume_sand_cushion = [(type_trench, lenght, volume_sand), (type_trench, lenght, volume_sand), ...]
    
#Транзакция
    with adoc.LockDocument():
        with adoc.Database as db:
            try:
                with db.TransactionManager.StartTransaction() as t:
                    #Точка вставки в модель
                    ppo = PromptPointOptions('\nУкажите точку вставки таблицы с параметрами длин траншей и объема песка для выполнения подушки:')
                    pr = editor.GetPoint(ppo)
                    if pr.Status != PromptStatus.OK:
                        editor.WriteMessage('\nВыполнение прервано')
                        t.Commit()
                        return
                    point = pr.Value #окончание функции точки вставки
                    #Создание названий столбцов
                    dict_col = { #наименования столбцов
                        TYPE_TR: 0,
                        LENGTH_ALIG: 1,
                        SAND_CUSHION: 2,
                        SAND_BACKFILL: 3,
                        SOIL_BACKFILL: 4,
                        }
                    header = 2  #количество строк в шапке таблицы
                    acad_table = AcadTable() #функция создания таблицы
                    acad_table.NumRows = header + len(data) #количество строк во всей таблице
                    acad_table.NumColumns = len(dict_col) #указание количества столбцов
                    acad_table.Position = point #указание для таблицы точки вставки
                    
                    #Создание шапки (Наименование)
                    acad_table.SetTextString(1, 0, f"{name_trench_style}") #наименование заголовка таблицы
                    acad_table.SetRowHeight(1, 6.5) # высота строки (№ - номер строки; № - высота строки)
                    acad_table.SetTextHeight(1, 0, 2.5) # высота текста в ячейке (№строки, № столбца, № высота кегля)
                    c_range = CellRange.Create(acad_table, 1, dict_col[TYPE_TR], 1, dict_col[SOIL_BACKFILL]) #создание диапозона ячейки для объединения (заголовок в данном случае) - №строки первой ячейки;название первого столбца; №строки конечной ячейки; название конечного столбца
                    acad_table.MergeCells(c_range) # функция объединения
                    acad_table.SetAlignment(1, 0, CellAlignment.MiddleCenter) #центрирование внутри ячейки
                    #Создание шапки (Основные столбцы)
                    acad_table.SetRowHeight(0, 10) #указание высоты строки (№ - номер строки; № - высота строки)
                    acad_table.SetTextString(0, dict_col[TYPE_TR], TYPE_TR) #занести текст в ячейки (названия столбцов)
                    acad_table.SetTextString(0, dict_col[LENGTH_ALIG], LENGTH_ALIG)
                    acad_table.SetTextString(0, dict_col[SAND_CUSHION], SAND_CUSHION)
                    acad_table.SetTextString(0, dict_col[SAND_BACKFILL], SAND_BACKFILL)
                    acad_table.SetTextString(0, dict_col[SOIL_BACKFILL], SOIL_BACKFILL)
    #Оформление таблицы (Ширина столбцов, высота текста и центрирование текста)
                    for row in range(header, header+len(data)):
                        acad_table.SetRowHeight(row, 6.5) #высота строк после заголовка
                    for col in range(len(dict_col)): #ширина столбцов от "начальной ячейки" до "конечной ячейки"
                        if col == dict_col[TYPE_TR]:
                            acad_table.SetColumnWidth(col, 25)
                        elif col == dict_col[LENGTH_ALIG]:
                            acad_table.SetColumnWidth(col, 25)
                        elif col == dict_col[SAND_CUSHION]:
                            acad_table.SetColumnWidth(col, 25)
                        elif col == dict_col[SAND_BACKFILL]:
                            acad_table.SetColumnWidth(col, 25)
                        elif col == dict_col[SOIL_BACKFILL]:
                            acad_table.SetColumnWidth(col, 25)
    
                    for row in range(0, header + len(data)): #высота текста во всех строках после шапки таблицы
                        for col in range(len(dict_col)): #цикл он начального до конечного столбца
                            acad_table.SetTextHeight(row, col, 3.5) #высота текста
                            acad_table.SetAlignment(row, col, CellAlignment.MiddleCenter) #выравнивание по середине
                            
                    #Наполнение данных таблицы (строка, столбец, значение)
                    row_table = header
                    for info_trench in data:
    
                        for col in range(len(dict_col)):
                            if col == dict_col[TYPE_TR]:
                                acad_table.SetTextString(row_table, col, f'{info_trench[0]}')
                            elif col == dict_col[LENGTH_ALIG]:
                                acad_table.SetTextString(row_table, col, f'{info_trench[1]}')
                            elif col == dict_col[SAND_CUSHION]:
                                acad_table.SetTextString(row_table, col, f'{info_trench[2]}')
                        
                        row_table += 1
    
                        
                        #Вставка таблицы в пространство модели
                    btr = t.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                    btr.AppendEntity(acad_table)
                    t.AddNewlyCreatedDBObject(acad_table, True)
                    t.Commit()
            except Exception as ex:
                editor.WriteMessage(f"\nКакая-то обшибка: \n {ex.Message}")
                    
        return data #было point


# функция оказалась лишней, т.к. трассы ЭН и ЭС находятся в разных шаблонах и на одной модели они не совпадают.
#def define_style(): # определяем какой стиль траншеи ищем (ЭН или ЭС)
#    with adoc.LockDocument():
#        with adoc.Database as db:
#    
#            with db.TransactionManager.StartTransaction() as t:
#                # выбор: присвоить номера трассе идущей от существующей опоры или от ШНО
#                opt = PromptKeywordOptions('\nРассчитать объём подушки для траншей ЭС или ЭН?')
#                opt.Keywords.Add('ЭС')
#                opt.Keywords.Add('ЭН')
#                des = editor.GetKeywords(opt)
#                if des.Status != PromptStatus.OK:
#                    editor.WriteMessage('\nВыполнение прервано')
#                    t.Commit()
#                    return
#                elif des.StringResult == 'ЭС':
#                    style_trench = TRENCH_STYLE_ES
#                else:
#                    style_trench = TRENCH_STYLE_EN
#    return style_trench

def find_alignments(): # нахождение трасс траншей
    with adoc.LockDocument():
        with adoc.Database as db:
            with db.TransactionManager.StartTransaction() as t:
                # находим все трассы на чертеже
                alignments = CivilDocument.GetSitelessAlignmentIds(civdoc)
                alignments_objects = [t.GetObject(i, OpenMode.ForRead) for i in alignments]
                alignments_filtered = [alignments_filtered for alignments_filtered in alignments_objects if TRENCH_STYLE in alignments_filtered.get_StyleName().lower()]
    return alignments_filtered


def check_name_trench(alignments): # проверяем на корреткность наименования трасс траншей
    with adoc.LockDocument():
        with adoc.Database as db:
            with db.TransactionManager.StartTransaction() as t:
                problem_trench = []
                normal_trench = []
                for i in range(len(alignments)):
                    if re.search(PATTERN_NAME_TRENCH, alignments[i].Name) == None:
                        problem_trench.append(alignments[i].Name)
                    else:
                        normal_trench.append(alignments[i])
                if problem_trench:
                    str = '\n'.join(problem_trench)
                    opt = PromptKeywordOptions(f'\nВ данном перечне трасс траншей некорректно заполнено наименование.\nДля выхода из скрипта нажмите "Esc".\nЕсли хотите продолжить без этих трасс, то нажмите "Да".\n{str}\n')
                    opt.Keywords.Add('Да')
                    des = editor.GetKeywords(opt)
                    if des.Status != PromptStatus.OK:
                        editor.WriteMessage('\nВыполнение прервано')
#                        t.Commit()
                        return
                    elif des.StringResult == 'Да':
                        return normal_trench
                else:
                    return normal_trench


def get_trench_style(temp_filtred_alignments):
    if TRENCH_STYLE_ES in temp_filtred_alignments[0].Name:
        return TRENCH_STYLE_ES
    else:
        return TRENCH_STYLE_EN

def check_type_trench(filtred_alignments): # проверяем на корреткность наименования типов траншей в НХ
    with adoc.LockDocument():
        with adoc.Database as db:
            with db.TransactionManager.StartTransaction() as t:
                problem_trench = []
                normal_trench = []
                dpsd = DictionaryPropertySetDefinitions(db)
                for i in range(len(dpsd.get_Records())):
                    if t.GetObject(dpsd.get_Records()[i], OpenMode.ForRead).Name == SET_CHARACTERISTICS:
                        alig_psd_id = t.GetObject(dpsd.get_Records()[i], OpenMode.ForRead).Id
                        break
                for alig in filtred_alignments: # перебираем трассы траншей для проверки имени характеристик
                    obj = alig # присваем переменной obj объект трассы траншеи
                    # назначение физировки в наборе характеристик
                    PropertyDataServices.AddPropertySet(obj, alig_psd_id)
                    alig_pds_id = PropertyDataServices.GetPropertySets(obj)
                    alig_pds = t.GetObject(alig_pds_id[0], OpenMode.ForRead)
                    prop_sets_data = alig_pds.PropertySetData

                    for psdata in prop_sets_data:
                        if psdata.Id == alig_pds.PropertyNameToId(TRENCH_TYPE): # ищем совпадение значения строки из НХ с необходимым нам 'Тип траншеи'
                            type_trench = alig_pds.GetAt(alig_pds.PropertyNameToId(TRENCH_TYPE))
                            if re.search(PATTERN_TYPE_TRENCH, type_trench) == None:
                                problem_trench.append(alig.Name)
                            else:
                                width_trench = alig_pds.GetAt(alig_pds.PropertyNameToId(TRENCH_WIDTH)) # ширина траншеи
                                height_trench = alig_pds.GetAt(alig_pds.PropertyNameToId(TRENCH_HEIGHT)) # выоста траншеи
                                sand_backfill = alig_pds.GetAt(alig_pds.PropertyNameToId(TRENCH_SAND_BACKFILL)) # высота песчанной засыпки
                                sand_cushion = alig_pds.GetAt(alig_pds.PropertyNameToId(TRENCH_SAND_CUSHION)) # высота песчанной подушки
                                slope_steepness = alig_pds.GetAt(alig_pds.PropertyNameToId(TRENCH_SLOPE_STEEPNESS)) # крутизна откоса
                                
                                result_length = 0 # общая длина трассы
                                for piece in alig.Entities:
                                    result_length += piece.get_Length()
                                normal_trench.append((type_trench, height_trench, width_trench, sand_cushion, sand_backfill, alig.Name, result_length, slope_steepness))

                if problem_trench:
                    str = '\n'.join(problem_trench)
                    opt = PromptKeywordOptions(f'\nВ данном перечне трасс траншей некорректно заполнены типы траншей в НХ.\nДля выхода из скрипта нажмите "Esc".\nЕсли хотите продолжить без этих трасс, то нажмите "Да".\n{str}\n')
                    opt.Keywords.Add('Да')
                    des = editor.GetKeywords(opt)
                    if des.Status != PromptStatus.OK:
                        editor.WriteMessage('\nВыполнение прервано')
#                        t.Commit()
                        return
                    elif des.StringResult == 'Да':
                        return normal_trench
                else:
                    return normal_trench


def create_dict_trench_types(alignments): # формируем словарь из ключ - типы траншей, значение именна трасс с этими типами траншей
    with adoc.LockDocument():
        with adoc.Database as db:
            with db.TransactionManager.StartTransaction() as t:
                types_trench_dict = {} # создаем словарь
                for alig in alignments:
                    types_trench_dict.setdefault((alig[0]), []).append((alig[5], alig[6], alig[1], alig[2], alig[3], alig[4], alig[7])) # вставляем для нашего ключа (тип траншеи) трассу относящуюся к типу траншеи с параметрами (имя трассы, длина, высота, ширина, высота песчанной подушки, высота песчанной засыпки, крутизна откоса)

                return types_trench_dict


def check_match_type_alig(aligs_typle):  # проверяем соответствие имени типа траншеи с именем трассы траншеи
    with adoc.LockDocument():
        with adoc.Database as db:
            with db.TransactionManager.StartTransaction() as t:
                problem_trench = []
                normal_trench = []
                for alig in aligs_typle:
                    trench_type = alig[0] # присваем тим траншеи
                    count_point = alig[5].count(".", alig[5].rfind("_")) # ищем количество точек в имени трассы после самого правого симовола _. Нужно чтобы корректно получить имя траншеи из имени трассы.
                    if count_point == 1:
                        name_trench_alig = alig[5][alig[5].rfind("_")+1:alig[5].rfind(".")] # получили тип траншеи из имени трассы, если в нашем случае есть точка, по типу ТК-4.1
                    else:
                        name_trench_alig = alig[5][alig[5].rfind("_")+1:] # получили тип траншеи из имени трассы, если в нашем случае нет точки, по типу ТК-4
                    
                    if trench_type == name_trench_alig:
                        normal_trench.append(alig)
                    else:
                        problem_trench.append(alig[5])
                if problem_trench:
                    str = '\n'.join(problem_trench)
                    opt = PromptKeywordOptions(f'\nВ данном перечне указаны трассы траншей которые не соответствуют именам в наборах характеристик.\nНеобходимо привести данные в соответствие.\nДля выхода из скрипта нажмите "Esc".\nЕсли хотите продолжить без этих трасс, то нажмите "Да".\n{str}\n')
                    opt.Keywords.Add('Да')
                    des = editor.GetKeywords(opt)
                    if des.Status != PromptStatus.OK:
                        editor.WriteMessage('\nВыполнение прервано')
#                        t.Commit()
                        return
                    elif des.StringResult == 'Да':
                        return normal_trench
                else:
                    return normal_trench


def check_parametrs_aligs(temp_trench_dict): # проверяем на соответствие параметров НХ для каждого типа траншей по трассам относящихся к данному типу траншей
    trench_dict = {}
    problem_trench = []
    for key, aligs in temp_trench_dict.items():
        if len(aligs) == 1:
            length_trench = aligs[0][1] # длина траншеи
            height_trench = aligs[0][2] # высота траншеи
            width_trench = aligs[0][3] # ширина траншеи
            sand_cushion = aligs[0][4] # высота песчанной подушки
            sand_backfill = aligs[0][5] # высота песчанной засыпки
            slope_steepness = aligs[0][6] # крутизна откоса
            trench_dict[key] = (length_trench, width_trench, sand_cushion, int(slope_steepness)) # составляем взаимосвязь ключ-значение.! ВНИМАНИЕ! нужно крутизне откоса назначить числовое значение int
            
        else:
            length_trench = aligs[0][1] # длина траншеи начальная
            # параметры ниже которые будут сравниваться по каждой трассе данного типа траншеи
            et_alig_name = aligs[0][0] # имя траншеи эталонной
            et_height_trench = aligs[0][2] # высота траншеи эталонной
            et_width_trench = aligs[0][3] # ширина траншеи эталонной
            et_sand_cushion = aligs[0][4] # высота песчанной подушки эталонной
            et_sand_backfill = aligs[0][5] # высота песчанной засыпки эталонной
            et_slope_steepness = aligs[0][6] # крутизна откоса текущей
            et_type_trench = [et_height_trench, et_width_trench, et_sand_cushion, et_sand_backfill, et_slope_steepness]
            
            for i in range(1, len(aligs)):
                alig_name = aligs[i][0] # имя траншеи текущей
                height_trench = aligs[i][2] # высота траншеи текущей
                width_trench = aligs[i][3] # ширина траншеи текущей
                sand_cushion = aligs[i][4] # высота песчанной подушки текущей
                sand_backfill = aligs[i][5] # высота песчанной засыпки текущей
                slope_steepness = aligs[i][6] # крутизна откоса текущей
                type_trench = [height_trench, width_trench, sand_cushion, sand_backfill, slope_steepness]
                if et_type_trench != type_trench: # если не соответствует собираем данный в проблемный список
                    problem_trench.append(f"{et_alig_name} !!!{alig_name}!!!")
                else: # если всё хорошо то добалвяем длину текущей трассы к общей длине данного типа траншей
                    length_trench += aligs[i][1]
            trench_dict[key] = (length_trench, et_width_trench, et_sand_cushion, int(slope_steepness))

    if problem_trench:
        str = '\n'.join(problem_trench)
        opt = PromptKeywordOptions(f'\nВ данном перечне указаны трассы у которых не совпадают один или несколько параметров НХ.\nНеобходимо привести данные в соответствие.\nДля выхода из скрипта нажмите "Esc".\nЕсли хотите продолжить без этих трасс (выделены символами !!!), то нажмите "Да".\n{str}\n')
        opt.Keywords.Add('Да')
        des = editor.GetKeywords(opt)
        if des.Status != PromptStatus.OK:
            editor.WriteMessage('\nВыполнение прервано')
#            t.Commit()
            return
        elif des.StringResult == 'Да':
            return trench_dict
    else:
        return trench_dict


def get_volume_sand_cushion(trench_dict): # получим тип-траншеи - длина, объём песчанной подушки
    volume_sand_cushion = []
    for key, info in trench_dict.items(): # key - тип траншеи, info[0] - длина траншеи, info[1] - ширина траншеи, info[2] - высота подушки, info[3] - крутизна откоса
        volume_sand_cushion.append((key, round(info[0], 1), round(info[0] * (info[1] * info[2] + (info[2] * 1 / math.tan(math.radians(info[3])) * info[2])) * KF_SAND_COMPACTION, 1)))

    # сортируем список по возрастанию нумерации типа траншеи (ТК-1, ТК-2 и т.д.)
    volume_sand_cushion.sort(key=lambda x: int(x[0][x[0].find('-')+1:]))
    return volume_sand_cushion


def main():
    with adoc.LockDocument():
        with adoc.Database as db:
            with db.TransactionManager.StartTransaction() as t:
#                style_trench = define_style()
                alignments = find_alignments() # находим нужные нам трассы траншей
                temp_filtred_alignments = check_name_trench(alignments) # проверяем трассы траншей на соответствие имени трасс необходимоу шаблону
                if temp_filtred_alignments is None:
                    return
                name_trench_style = get_trench_style(temp_filtred_alignments) # получим стиль траншеи ЭН или ЭС. Потом нужно для составления таблицы
                temp_filtred_alignments_typle = check_type_trench(temp_filtred_alignments) # проверяем трассы траншей на соответствие имени типа трншей в НХ  необходимоу шаблону
                if temp_filtred_alignments_typle is None:
                    return
                filtred_alignments_typle = check_match_type_alig(temp_filtred_alignments_typle) # проверяем соответствие имени типа траншеи с именем трассы траншеи
                if filtred_alignments_typle is None:
                    return
                temp_trench_dict = create_dict_trench_types(filtred_alignments_typle) # формируем словарь из ключ - типы траншей, значение именна трасс с этими типами траншей и плюс параметры
                if temp_trench_dict is None:
                    return
                trench_dict = check_parametrs_aligs(temp_trench_dict) # проверяем на соответствие параметров НХ для каждого типа траншей по трассам относящихся к данному типу траншей
                if trench_dict is None:
                    return
                volume_sand_cushion = get_volume_sand_cushion(trench_dict) # получим тип-траншеи - длина, объём песчанной подушки
                data = create_table(volume_sand_cushion, name_trench_style)
    
                    

                # Commit before end transaction
                t.Commit() 
#    return trench_dict
    return [volume_sand_cushion, trench_dict, data]


# Assign your output to the OUT variable.
OUT = main()
