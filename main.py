import os
import csv
import doctest
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH


class GradeSheet():
    """
    >>> gs = GradeSheet()
    >>> gs.getFileCSV('E://volon/files/v.csv')
    'file exists'
    >>> gs._pathFileCSV
    'E://volon/files/v.csv'
    >>> gs.getDict()
    {...}
    """

    _pathFileCSV = ''
    _dirForSave = ''

    def getFileCSV(self, path):
        if os.path.exists(path):
            self._pathFileCSV = path
            return 'file exists'
        return 'file not exists'

    def getDirForSave(self, path):
        if os.path.isdir(path):
            self._dirForSave = path
            return 'path is not dir'
        return 'path is dir'

    def _getDict(self):
        if self._pathFileCSV == '':
            return 'path is empty'
        dict_person = {}

        with open(self._pathFileCSV, newline='') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            for row in reader:
                if row[18] == '1':
                    if dict_person.get(row[12]):
                        if dict_person.get(row[12]).get(row[11]):
                            dict_person[row[12]][row[11]] += 1
                        else:
                            dict_person[row[12]][row[11]] = 1
                    else:
                        dict_person[row[12]] = {row[11]: 1}

        return dict_person

        # for i in dict_person:
        #     for k in dict_person[i]:
        #         print(i, k, dict_person[i][k])

    def createDocs(self, dict_person):

        for faculty in dict_person:
            # Создаем документ
            doc = docx.Document()

            # Заголовок
            doc.add_heading('ОЦЕНОЧНАЯ ВЕДОМОСТЬ ПО ПРОЕКТУ', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            # добавляем параграфы
            doc.add_paragraph('Волонтеры: олимпиадный марафон (название проекта)').alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph('Сервисный проект (тип проекта)').alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph('Январь – май (срок выполнения проекта)').alignment = WD_ALIGN_PARAGRAPH.CENTER

            # добавляем пурвую таблицу 2х2
            table = doc.add_table(rows=2, cols=2)
            # выравниваем ее по левому краю
            table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            # настраиваем ширину столбцов (через ячейки, т.к. word не равняет ширину колонок через column)
            cells0 = table.columns[0].cells
            for i in cells0:
                i.width = 373224 * 5
            cells1 = table.columns[1].cells
            for i in cells1:
                i.width = 373224 * 7
            # применяем стиль для таблицы
            table.style = 'Table Grid'
            # заполняем таблицу
            one = table.cell(0, 0)
            one.text = ''
            one.paragraphs[0].add_run('Руководитель проекта: ').bold = True
            one.add_paragraph('ФИО ')
            one.add_paragraph('Должность ')
            two = table.cell(0, 1)
            two.add_paragraph().add_run('Протасевич Тамара Анатольевна').bold = True
            two.add_paragraph('Директор по профессиональной ориентации и работе с одаренными учащимися')
            table.cell(1, 0).text = 'Образовательная программа'
            table.cell(1, 1).text = faculty

            # пропускаем строчку
            doc.add_paragraph('')

            # добавляем таблицу 1x4
            tableTwo = doc.add_table(rows=1, cols=4)
            # применяем стиль для таблицы
            tableTwo.style = 'Table Grid'
            # заполняем заголовки для страницы
            tableTwo.cell(0, 0).text = 'ФИО'
            tableTwo.cell(0, 1).text = 'группа'
            tableTwo.cell(0, 2).text = 'Оценка по 10-балльной шкале'
            tableTwo.cell(0, 3).text = 'Количество ЗЕ за проект'

            # добавляем строки таблице и заполняем содержимым переданного словарая с волонтерами
            for person in dict_person[faculty]:
                r = tableTwo.add_row()
                c = r.cells
                c[0].text = person
                c[2].text = '10'
                grade = dict_person[faculty][person]
                if grade < 3:
                    value = '1'
                elif grade == 3:
                    value = '2'
                else:
                    value = '3'
                c[3].text = value

            # задаем ширину столбцов, где 373224 количество EMU в одном сантиметре
            cells0 = tableTwo.columns[0].cells
            for i in cells0:
                i.width = 373224 * 8
            cells1 = tableTwo.columns[1].cells
            for i in cells1:
                i.width = 373224 * 2

            # путь сохранения файла
            path_save = faculty.replace('"', '').replace(':', '')
            a = f'{self._dirForSave}/{path_save}.docx'
            # print(a) # тестовый вывод

            # сохраняем созданный документ
            doc.save(a)


if __name__ == '__main__':
    doctest.testmod(optionflags=+doctest.ELLIPSIS)
    # gs = GradeSheet()
    # gs._dirForSave = 'f'
    # gs.getFileCSV('E://volon/files/v.csv')
    # gs.createDocs(gs._getDict())
