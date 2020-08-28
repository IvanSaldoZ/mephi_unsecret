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
        self.unsecret_department_input_file_path = os.path.join(self.input_forms_folder_name, 'unsecret_department.rtf')
        self.unsecret_identification_input_file_path = os.path.join(self.input_forms_folder_name, 'unsecret_identification.doc')
        self.unsecret_university_input_file_path = os.path.join(self.input_forms_folder_name, 'unsecret_university.docx')
        self.unsecret_department_output_file_path = os.path.join(self.output_forms_folder_name, 'unsecret_department.rtf')
        self.unsecret_identification_output_file_path = os.path.join(self.output_forms_folder_name, 'unsecret_identification.doc')
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
        self.replace_university(input_params)

    def replace_university(self, input_params):
        """Заменяем в файле рассекречивания университета нужные поля"""
        context = {
            'object_name': input_params['object_name']
        }
        self.replace_in_docx(self.unsecret_university_input_file_path,
                             context,
                             self.unsecret_university_output_file_path)

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