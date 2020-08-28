import configparser  # For ini-reading
import os
from docxtpl import DocxTemplate



class UnsecretMephiProg:
    """Основной класс программы для формирования файлов на рассекречивание"""

    def __init__(self):
        self.__input_file_name: str = 'input.txt'
        self.__tmp_file_name: str = 'input.tmp'
        self.input_forms_folder_name: str = 'forms'
        self.output_forms_folder_name: str = 'output'
        self.unsecret_department_input_file_path = os.path.join(self.input_forms_folder_name, 'unsecret_department.docx')
        self.unsecret_identification_input_file_path = os.path.join(self.input_forms_folder_name, 'unsecret_identification.docx')
        self.unsecret_university_input_file_path = os.path.join(self.input_forms_folder_name, 'unsecret_university.docx')
        self.unsecret_department_output_file_path = os.path.join(self.output_forms_folder_name, 'unsecret_department.docx')
        self.unsecret_identification_output_file_path = os.path.join(self.output_forms_folder_name, 'unsecret_identification.docx')
        self.unsecret_university_output_file_path = os.path.join(self.output_forms_folder_name, 'unsecret_university.docx')

    def read_input_file(self):
        """
        Метод для чтения входного файла .dat в формате ini с моделью кориума с комментариями
        :return:
        """
        self.remove_comments()  # Сначала удаляем комменты из файла
        """Читаем входной файл и возвращаем его содержимое"""
        inputfile = configparser.ConfigParser()
        try:
            inputfile.read(self.__tmp_file_name, encoding='utf-8')
        except Exception as e:
            print('Error writing to the folder of the program. Please change the permissions to the program folder')
            raise e
        self.remove_tmpfile()
        return inputfile  # -> возвращаем содержимое INI-файла

    def remove_comments(self, separator='#'):  # Remove comments from the file
        """
        Метод для удаления комментов из входного файла (после символа #)
        """
        try:
            f1 = open(self.__input_file_name, "r", encoding='utf-8')
            f2 = open(self.__tmp_file_name, "w", encoding='utf-8')
            for line in f1.readlines():
                # Remove comment from the line and save result to temp file
                f2.write(line.split(separator)[0] + '\n')
            f1.close()
            res = f2
            f2.close()
        except Exception as e:
            print('Input {self.__input_file_name} file not found')
            raise e
        return res

    def remove_tmpfile(self) -> None:
        """
        Remove temporary file
        :return:
        """
        try:
            os.remove(self.__tmp_file_name)
        except Exception as e:
            print('Cannot remove temporary file')
            raise e

    def replace_all(self, input_params):
        """Запускаем методы для замены каждого типа форм"""
        self.replace_university(input_params)
        self.replace_deparment(input_params)
        self.replace_identification(input_params)

    def replace_university(self, context: dict):
        """Заменяем в файле рассекречивания университета нужные поля"""
        self.replace_in_docx(self.unsecret_university_input_file_path,
                             context,
                             self.unsecret_university_output_file_path)

    def replace_deparment(self, context: dict):
        """Заменяем в файле рассекречивания департамента нужные поля"""
        self.replace_in_docx(self.unsecret_department_input_file_path,
                             context,
                             self.unsecret_department_output_file_path)

    def replace_identification(self, context: dict):
        """Заменяем в файле идентификационного заключения нужные поля"""
        self.replace_in_docx(self.unsecret_identification_input_file_path,
                             context,
                             self.unsecret_identification_output_file_path)

    def replace_in_docx(self, file_name: str, context_input: dict, output_file_name: str):
        """Заменяем в docx-файле file_name нужные поля, заданные в context_input, и сохраняем результат
        в output_file_name"""
        doc = DocxTemplate(file_name)
        doc.render(context_input)
        doc.save(output_file_name)

    def transfer_ini_to_dict(self, input_params_ini):
        """Преобразуем входные параметры в представление dict, чтобы если изменились параметры входного файла,
        внутреннее представление оставалось неизмеенным и всё работало, не зависимо от формата входного файла"""
        input_params_dict = {}
        input_params_dict['object_name'] = input_params_ini["GENERAL"]['object_name']
        input_params_dict['authors_rod'] = input_params_ini["GENERAL"]['authors_rod']
        input_params_dict['material_name_im'] = input_params_ini["GENERAL"]['material_name_im']
        input_params_dict['material_name_rod'] = input_params_ini["GENERAL"]['material_name_rod']
        input_params_dict['to_where'] = input_params_ini["GENERAL"]['to_where']
        input_params_dict['material_name_im'] = input_params_ini["GENERAL"]['material_name_im']
        input_params_dict['material_name_rod'] = input_params_ini["GENERAL"]['material_name_rod']
        input_params_dict['ih'] = input_params_ini["GENERAL"]['ih']
        input_params_dict['nih'] = input_params_ini["GENERAL"]['nih']
        input_params_dict['prepared_im'] = input_params_ini["GENERAL"]['prepared_im']
        input_params_dict['prepared_rod'] = input_params_ini["GENERAL"]['prepared_rod']
        input_params_dict['committee_title_1'] = input_params_ini["GENERAL"]['committee_title_1']
        input_params_dict['committee_title_2'] = input_params_ini["GENERAL"]['committee_title_2']
        input_params_dict['committee_name'] = input_params_ini["GENERAL"]['committee_name']
        input_params_dict['member_1'] = input_params_ini["GENERAL"]['member_1']
        input_params_dict['member_2'] = input_params_ini["GENERAL"]['member_2']
        input_params_dict['member_3'] = input_params_ini["GENERAL"]['member_3']
        input_params_dict['member_4'] = input_params_ini["GENERAL"]['member_4']
        input_params_dict['member_5'] = input_params_ini["GENERAL"]['member_5']
        input_params_dict['member_6'] = input_params_ini["GENERAL"]['member_6']
        input_params_dict['member_7'] = input_params_ini["GENERAL"]['member_7']
        input_params_dict['member_8'] = input_params_ini["GENERAL"]['member_8']
        input_params_dict['commission_number'] = input_params_ini["GENERAL"]['commission_number']
        input_params_dict['commission_name'] = input_params_ini["GENERAL"]['commission_name']
        input_params_dict['buyer'] = input_params_ini["GENERAL"]['buyer']
        input_params_dict['annotation'] = input_params_ini["GENERAL"]['annotation']
        input_params_dict['key_words'] = input_params_ini["GENERAL"]['key_words']
        input_params_dict['conclusion'] = input_params_ini["GENERAL"]['conclusion']
        input_params_dict['material_name_im_capital'] = input_params_ini["GENERAL"]['material_name_im'].capitalize()

        return input_params_dict


    def run(self):
        """Главная программа"""
        # Считываем файл с входными параметрами матриалов доклада
        input_params_ini = self.read_input_file()
        # Преобразуем в dict
        input_params_dict = self.transfer_ini_to_dict(input_params_ini)
        # Заменяем необходимые поля во всех формах
        self.replace_all(input_params_dict)
        print('Done!')


if __name__ == "__main__":
    prog = UnsecretMephiProg()
    prog.run() # Запускаем