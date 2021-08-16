from config import *
import traceback


class DOCUMENT():
    def __init__(self, obligation, file_name, folder_name, obligation_b):
        self.obligation = obligation
        self.obligation_b = obligation_b
        self.file_name = file_name
        self.folder_name = folder_name
        self.form_metadata = []
        if len( self.file_name.split('_') ) == 1:
            self.form_metadata.append('Название документа')
            self.document_name = self.file_name
        else:
            for self.a in self.file_name.split('_'):
                self.form_metadata.append(self.a)
            self.document_name = self.file_name.split('_')[1]

    def pass_metadata(self, non_split_name):
        self.dict_metadata = {}
        self.split_name = non_split_name.split('_')
        self.i = 0
        try:
            for self.a in self.form_metadata:
                if self.a == 'Дата':
                    self.redact_date = self.split_name[self.i][6:8]+'.'\
                                    +self.split_name[self.i][4:6]+'.'\
                                    +self.split_name[self.i][0:4]
                    self.dict_metadata[self.a] = self.redact_date
                    self.i += 1
                    continue
                if self.i == 1:
                    self.dict_metadata['Название документа'] = self.split_name[self.i]
                    self.i += 1
                    continue
                self.dict_metadata[self.a] = self.split_name[self.i]
                self.i += 1
        except:
            print('Ошибка формата')
            self.dict_metadata = {}
            self.dict_metadata['Название документа'] = non_split_name
        return self.dict_metadata

    def check_bum_v(self, non_split_name):
        self.split_name = non_split_name.split('_')
        try:
            if self.split_name[len(self.form_metadata)] == 'Б':
                return True
            else:
                return False
        except:
            return False

class KONTRAGENT():
    def __init__(self, document_models, kontragent_name):
        self.document_models = document_models
        self.kontragent_name = kontragent_name

    def check(self):
        """
        Проверка всех ключей для Зоны 2

        :return:
        """
        self.dict_result_indicators = {}
        for self.a_ch, self.b_ch in self.document_models.items():
            self.dict_result_indicators[self.a_ch] = self.check_presence_file(model=self.b_ch)
        return self.dict_result_indicators

    def check_presence_file(self, model):
        self.presence_file_list = []
        self.path = DIR_HRANILIHE + self.kontragent_name + '/' + model.folder_name + '/'
        print(self.path)
        self.search_words = model.document_name

        self.bum_v = False
        self.result_check = False
        try:
            self.list_files = os.listdir(self.path)
            self.list_files_wo_r = []
            for self.a in self.list_files:
                if self.a.find('.pdf') != -1:
                    self.list_files_wo_r.append(self.a.split('.')[0])

            for self.file in self.list_files_wo_r:
                self.bum_v = False
                self.result_check = False

                if self.file.find('_') != -1:
                    if self.file.split('_')[1] == self.search_words:
                        self.result_check = True
                        self.bum_v = model.check_bum_v(self.file)
                        self.presence_file_list.append([self.result_check, self.bum_v, self.file])
                else:
                    if self.file == self.search_words:
                        self.result_check = True
                        self.bum_v = False
                        self.presence_file_list.append([self.result_check, self.bum_v, self.file])
        except Exception as err:
            self.result_check = False
            self.bum_v = False
            print('Нет папки: ', traceback.format_exc())


        if self.presence_file_list == []:
            return [[self.result_check, self.bum_v, model.document_name]]
        else:
            return self.presence_file_list

    def check_presence_file_w_filename(self, model):
        """
        Для таблицы эксель. С именем файла

        :param model:
        :return:
        """
        self.presence_file_list = []
        self.path = DIR_HRANILIHE + '/' + self.kontragent_name + '/' + model.folder_name + '/'
        self.search_words = model.document_name
        try:
            self.list_files = os.listdir(self.path)
            self.list_files_wo_r = []
            for self.a in self.list_files:
                if self.a.find('.pdf') != -1:
                    self.list_files_wo_r.append(self.a.split('.')[0])

            self.file_name_check_presence = ''
            # self.bum_v = False
            # self.result_check = False
            for self.file in self.list_files_wo_r:
                if self.file.find('_') != -1:
                    if self.file.split('_')[1] == self.search_words:
                        self.file_name_check_presence = self.file
                        self.presence_file_list.append(self.file_name_check_presence)
                else:
                    if self.file == self.search_words:
                        self.file_name_check_presence = self.file
                        self.presence_file_list.append(self.file_name_check_presence)
        except Exception as err:
            print('Нет папки: ', err)

        return self.presence_file_list

    def check_w_filename(self):
        self.dict_result_check = {}
        for self.a_ch, self.b_ch in self.document_models.items():
            self.dict_result_check[self.a_ch] = self.check_presence_file_w_filename(model=self.b_ch)
        return self.dict_result_check