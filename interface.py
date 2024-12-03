from tkinter import *
from tkinter import messagebox, filedialog, ttk
from multiprocessing import Process, Queue, freeze_support
import config
import service


class MainWindow:
    def __init__(self) -> None:
        self.master = Tk()
        self.mpqueue = Queue()
        self.subprocess1 = None
        self.name_program = 'SFS DocRename'
        self.master.title(self.name_program)
        self.master.geometry('600x800')
        self.master.iconbitmap('microsoft-office-word.ico')
        self.master.protocol('WM_DELETE_WINDOW', self.__exit)

        # главное меню
        self.main_menu = Frame(self.master, relief=RAISED, bd=2)
        self.main_menu.pack(side=TOP, expand=NO, fill=X)
        self.menu_btn_about = Menubutton(self.main_menu,
                                         text='Menu',
                                         underline=0)
        self.menu_btn_about.pack(side=LEFT)
        self.menu = Menu(self.menu_btn_about, tearoff=0)
        self.menu.add_command(label='About', command=self.__show_about)
        self.menu_btn_about.configure(menu=self.menu)

        # фрейм верхнего уровня
        self.frame_top = Frame(self.master)
        self.frame_top.pack(side=TOP, fill=BOTH, expand=YES)

        # фрейм пути к папке с файлами
        self.frame_path = Frame(self.frame_top)
        self.frame_path.pack(side=TOP, fill=BOTH, expand=NO)

        self.frame_path_entry = LabelFrame(self.frame_path,
                                           bd=3,
                                           font=('arial', config.FONT_SIZE),
                                           text='Путь к файлам:')
        self.frame_path_entry.pack(side=LEFT, fill=X, expand=YES)
        self.select_path_button = Button(self.frame_path_entry,
                                         bd=2,
                                         text='Выбрать',
                                         font=('arial', config.FONT_SIZE),
                                         command=self.__select_file_dialog)
        self.select_path_button.pack(side=LEFT, pady=10)
        self.input_path = Entry(self.frame_path_entry,
                                bd=3,
                                width=60,
                                font=('arial', config.FONT_SIZE))
        self.input_path.pack(side=LEFT, pady=10, padx=10)

        # начало фрейма дерева ключевых слов
        self.frame_keywords_top = Frame(self.frame_top)
        self.frame_keywords_top.pack(side=TOP, fill=BOTH, expand=YES)

        self.frame_keywords_label = LabelFrame(self.frame_keywords_top,
                                               bd=3,
                                               font=(
                                                   'arial', config.FONT_SIZE),
                                               text='Фразы поиска:')
        self.frame_keywords_label.pack(side=LEFT, fill=BOTH, expand=YES)

        self.keywords = ttk.Treeview(self.frame_keywords_label)
        self.keywords.heading('#0', text='Список ключевых фраз', anchor=W)
        self.keywords.pack(side=TOP, fill=BOTH, expand=YES)
        self.keywords.bind('<<TreeviewSelect>>', self.__select_item_tree)
        self.keywords_scrollbar = ttk.Scrollbar(self.keywords,
                                                orient=VERTICAL,
                                                command=self.keywords.yview)
        self.keywords.configure(yscrollcommand=self.keywords_scrollbar.set)
        self.keywords_scrollbar.pack(side=RIGHT, fill=Y)

        self.frame_keywords_movies = Frame(self.frame_keywords_top)
        self.frame_keywords_movies.pack(side=RIGHT, fill=X, expand=NO)

        self.button_move_word_up = Button(self.frame_keywords_movies,
                                          bd=2,
                                          text='▲',
                                          font=('arial', config.FONT_SIZE),
                                          state=DISABLED,
                                          command=self.__move_word_up)
        self.button_move_word_up.pack(side=TOP, pady=10, padx=10)
        self.button_move_word_down = Button(self.frame_keywords_movies,
                                            bd=2,
                                            text='▼',
                                            font=('arial', config.FONT_SIZE),
                                            state=DISABLED,
                                            command=self.__move_word_down)
        self.button_move_word_down.pack(side=TOP, pady=10, padx=10)
        # Конец фрейма дерева ключевых слов

        self.frame_buttons = Frame(self.frame_top)
        self.frame_buttons.pack(side=TOP, fill=BOTH, expand=NO)

        self.button_add_keyword = Button(self.frame_buttons,
                                         bd=2,
                                         text='Добавить основную фразу',
                                         font=('arial', config.FONT_SIZE),
                                         command=self.__call_add_keyword)
        self.button_add_keyword.pack(side=LEFT, pady=10, padx=10)
        self.button_add_subword = Button(self.frame_buttons,
                                         bd=2,
                                         text='Добавить дополнительную фразу',
                                         font=('arial', config.FONT_SIZE),
                                         state=DISABLED,
                                         command=self.__call_add_subword)
        self.button_add_subword.pack(side=LEFT, pady=10, padx=10)
        self.button_delete_word = Button(self.frame_buttons,
                                         bd=2,
                                         text='Удалить фразу',
                                         font=('arial', config.FONT_SIZE),
                                         state=DISABLED,
                                         command=self.__delete_word)
        self.button_delete_word.pack(side=LEFT, pady=10, padx=10)

        self.frame_stop_words = Frame(self.frame_top)
        self.frame_stop_words.pack(side=TOP, fill=BOTH, expand=YES)

        self.frame_stop_words_label = LabelFrame(self.frame_stop_words,
                                                 bd=3,
                                                 font=(
                                                     'arial', config.FONT_SIZE),
                                                 text='Блокировать фразы в названии:')
        self.frame_stop_words_label.pack(side=TOP, fill=BOTH, expand=YES)

        self.listbox_stop_words = Listbox(self.frame_stop_words_label,
                                          bd=2,
                                          font=('arial', config.FONT_SIZE))
        self.listbox_stop_words.pack(side=TOP, fill=BOTH, expand=YES)
        self.listbox_stop_words.bind("<<ListboxSelect>>",
                                     self.__select_item_listbox)
        self.listbox_stop_words_scrollbar = ttk.Scrollbar(self.listbox_stop_words,
                                                          orient=VERTICAL,
                                                          command=self.listbox_stop_words.yview)
        self.listbox_stop_words.configure(
            yscrollcommand=self.listbox_stop_words_scrollbar.set)
        self.listbox_stop_words_scrollbar.pack(side=RIGHT, fill=Y)

        self.frame_stop_word_buttons = Frame(self.frame_stop_words)
        self.frame_stop_word_buttons.pack(side=TOP, fill=X, expand=NO)

        self.button_add_stop_word = Button(self.frame_stop_word_buttons,
                                           bd=2,
                                           text='Добавить фразу',
                                           font=('arial', config.FONT_SIZE),
                                           command=self.__call_add_stop_word)
        self.button_add_stop_word.pack(side=LEFT, pady=10, padx=10)
        self.button_delete_stop_word = Button(self.frame_stop_word_buttons,
                                              bd=2,
                                              text='Удалить фразу',
                                              font=('arial', config.FONT_SIZE),
                                              state=DISABLED,
                                              command=self.__delete_stop_word)
        self.button_delete_stop_word.pack(side=LEFT, pady=10, padx=10)

        self.frame_footer = Frame(self.frame_top)
        self.frame_footer.pack(side=TOP, fill=BOTH, expand=NO)

        self.min_count_symbols_in_doc_label = Label(self.frame_footer,
                                                    text='Минимальное количество символов в doc:')
        self.min_count_symbols_in_doc_label.pack(side=LEFT)
        self.min_count_symbols_in_doc = Entry(self.frame_footer,
                                              bd=3,
                                              width=10,
                                              font=('arial', config.FONT_SIZE))
        self.min_count_symbols_in_doc.pack(side=LEFT, pady=10, padx=10)
        self.button_start = Button(self.frame_footer,
                                   bd=3,
                                   text='Старт',
                                   font=('arial', config.FONT_SIZE),
                                   command=self.__start)
        self.button_start.pack(side=LEFT, pady=10, padx=10)

        # self.label_info = Label(self.frame_top, text='')
        self.label_info = Entry(self.frame_top,
                                bd=3,
                                width=60,
                                state='normal',
                                font=('arial', config.FONT_SIZE))
        self.label_info.pack(side=TOP, pady=1, padx=10, fill=BOTH)

        self.progressbar = ttk.Progressbar(self.frame_top,
                                           mode='determinate',
                                           length=600,
                                           value=0)
        self.progressbar.pack(side=BOTTOM, pady=10, padx=10, fill=BOTH)

        self.__load_config()
        self.master.mainloop()

    def __subprocess_close(self) -> None:
        if self.subprocess1:
            self.subprocess1.terminate()
            self.subprocess1.join()
            self.subprocess1.close()

    def __exit(self) -> None:
        self.__save()
        self.__subprocess_close()
        self.mpqueue.close()
        self.mpqueue.join_thread()
        self.master.destroy()

    def __save(self) -> None:
        path_directory = self.input_path.get()
        keywords = {}
        stop_words = []
        subwords = []
        min_symbols_in_doc = self.min_count_symbols_in_doc.get()

        if not min_symbols_in_doc.isdecimal():
            min_symbols_in_doc = service.get_default_min_symbols_in_doc()

        data_config = {'path_directory': path_directory,
                       'keywords': keywords,
                       'stop_words': stop_words,
                       'min_symbols_in_doc': min_symbols_in_doc}

        for parent_iid in self.keywords.get_children():
            item = self.keywords.item(parent_iid)
            keyword = item.get('text')

            for iid in self.keywords.get_children(parent_iid):
                item = self.keywords.item(iid)
                subword = item.get('text')
                subwords.append(subword)

            keywords[keyword] = subwords.copy()
            subwords.clear()

        stop_words.extend(self.listbox_stop_words.get(0, END))

        if keywords or stop_words:
            service.save_config(data_config)

    def __show_about(self) -> None:
        messagebox.showinfo(self.name_program, '"' + self.name_program +
                            '"' + ' powered by Hasan Smirnov(с) 2024')

    def __select_file_dialog(self) -> None:
        filepath = filedialog.askdirectory()

        if filepath:
            self.input_path.delete(0, END)
            self.input_path.insert(0, filepath)

    def __load_config(self) -> None:
        path_directory, keywords, stop_words, min_symbols_in_doc = service.load_config()

        if path_directory:
            self.input_path.insert(0, path_directory)

        iid = 0

        for key, words in keywords.items():
            self.keywords.insert(parent='', index=END, iid=iid, text=key)
            parent_iid = iid
            iid += 1

            for word in words:
                self.keywords.insert(parent=str(parent_iid),
                                     index=END, iid=iid, text=word)
                iid += 1

        for stop_word in stop_words:
            self.listbox_stop_words.insert(END, stop_word)

        if min_symbols_in_doc:
            self.min_count_symbols_in_doc.insert(0, min_symbols_in_doc)

    def __select_item_tree(self, event: Event) -> None:
        current_item = self.keywords.focus()

        if self.keywords.get_children(current_item):
            self.button_add_subword['state'] = ACTIVE
        else:
            self.button_add_subword['state'] = DISABLED

        self.button_delete_word['state'] = ACTIVE
        self.button_move_word_up['state'] = ACTIVE
        self.button_move_word_down['state'] = ACTIVE

    def __select_item_listbox(self, event: Event):
        self.button_delete_stop_word['state'] = ACTIVE

    def __call_add_keyword(self) -> None:
        add_text_modal = Toplevel(self.master)
        add_text_modal.iconbitmap('microsoft-office-word.ico')
        add_text_modal.title('Добавить основную фразу')
        input_text = Entry(add_text_modal,
                           bd=3,
                           width=60,
                           font=('arial', config.FONT_SIZE))
        input_text.pack(side=TOP, pady=10, padx=10)
        add_text_button = Button(add_text_modal,
                                 bd=2,
                                 text='Добавить',
                                 font=('arial', config.FONT_SIZE),
                                 command=lambda it=input_text, mw=add_text_modal: self.__add_keywords(it, mw))
        add_text_button.pack(side=LEFT, pady=10, padx=50)
        cancel_button = Button(add_text_modal,
                               bd=2,
                               text='Отменить',
                               font=('arial', config.FONT_SIZE),
                               command=add_text_modal.destroy)
        cancel_button.pack(side=RIGHT, pady=10, padx=50)
        input_text.focus_set()
        # удалю кнопки свернуть развернуть
        add_text_modal.transient(self.master)
        # передам поток модальному окну, что бы нельзя было переключить на главное окно
        add_text_modal.grab_set()
        # add_text_modal.wait_window()

    def __add_keywords(self, input_text: Entry, add_text_modal: Toplevel) -> None:
        keyword = input_text.get()
        self.keywords.insert(parent='', index=END, text=keyword)
        add_text_modal.destroy()

    def __call_add_subword(self) -> None:
        add_text_modal = Toplevel(self.master)
        add_text_modal.iconbitmap('microsoft-office-word.ico')
        add_text_modal.title('Добавить дополнительную фразу')
        input_text = Entry(add_text_modal,
                           bd=3,
                           width=60,
                           font=('arial', config.FONT_SIZE))
        input_text.pack(side=TOP, pady=10, padx=10)
        add_text_button = Button(add_text_modal,
                                 bd=2,
                                 text='Добавить',
                                 font=('arial', config.FONT_SIZE),
                                 command=lambda it=input_text, mw=add_text_modal: self.__add_subword(it, mw))
        add_text_button.pack(side=LEFT, pady=10, padx=50)
        cancel_button = Button(add_text_modal,
                               bd=2,
                               text='Отменить',
                               font=('arial', config.FONT_SIZE),
                               command=add_text_modal.destroy)
        cancel_button.pack(side=RIGHT, pady=10, padx=50)
        input_text.focus_set()
        add_text_modal.transient(self.master)
        add_text_modal.grab_set()

    def __add_subword(self, input_text: Entry, add_text_modal: Toplevel) -> None:
        current_item = self.keywords.focus()
        subword = input_text.get()
        self.keywords.insert(parent=current_item, index=END, text=subword)
        add_text_modal.destroy()

    def __delete_word(self) -> None:
        current_item = self.keywords.focus()
        self.keywords.delete(current_item)

    def __call_add_stop_word(self) -> None:
        add_text_modal = Toplevel(self.master)
        add_text_modal.iconbitmap('microsoft-office-word.ico')
        add_text_modal.title('Добавить фразу')
        input_text = Entry(add_text_modal,
                           bd=3,
                           width=60,
                           font=('arial', config.FONT_SIZE))
        input_text.pack(side=TOP, pady=10, padx=10)
        add_text_button = Button(add_text_modal,
                                 bd=2,
                                 text='Добавить',
                                 font=('arial', config.FONT_SIZE),
                                 command=lambda it=input_text, mw=add_text_modal: self.__add_stop_word(it, mw))
        add_text_button.pack(side=LEFT, pady=10, padx=50)
        cancel_button = Button(add_text_modal,
                               bd=2,
                               text='Отменить',
                               font=('arial', config.FONT_SIZE),
                               command=add_text_modal.destroy)
        cancel_button.pack(side=RIGHT, pady=10, padx=50)
        input_text.focus_set()
        add_text_modal.transient(self.master)
        add_text_modal.grab_set()

    def __add_stop_word(self, input_text: Entry, add_text_modal: Toplevel) -> None:
        stop_word = input_text.get()
        self.listbox_stop_words.insert(END, stop_word)
        add_text_modal.destroy()

    def __delete_stop_word(self) -> None:
        current_items = self.listbox_stop_words.curselection()

        for current_item in current_items:
            self.listbox_stop_words.delete(current_item)

    def __move_word_up(self) -> None:
        current_item = self.keywords.focus()
        top_level_keywords = self.keywords.get_children()
        parrent, low_level_keywords = self.__get_parrent_item(top_level_keywords,
                                                              current_item)

        if not parrent:
            idx = top_level_keywords.index(current_item)
        else:
            idx = low_level_keywords.index(current_item)

        new_idx = idx - 1

        if new_idx >= 0:
            self.keywords.move(current_item, parrent, new_idx)

    def __move_word_down(self) -> None:
        current_item = self.keywords.focus()
        top_level_keywords = self.keywords.get_children()
        parrent, low_level_keywords = self.__get_parrent_item(top_level_keywords,
                                                              current_item)

        if not parrent:
            idx = top_level_keywords.index(current_item)
        else:
            idx = low_level_keywords.index(current_item)

        new_idx = idx + 1
        self.keywords.move(current_item, parrent, new_idx)

    def __get_parrent_item(self, top_level_keywords: tuple, current_item: str) -> tuple[str, tuple]:
        parrent = ''
        low_level_keywords = tuple()

        # проверим принадлежит ли наш выделенный элемент какому либо родителю
        # если да то вернем id родителя и список остальных потомков
        for keyword in top_level_keywords:
            low_level_keywords = self.keywords.get_children(keyword)

            if low_level_keywords and current_item in low_level_keywords:
                parrent = keyword
                break

        return (parrent, low_level_keywords)

    def __get_info_from_subprocess(self) -> None:
        while not self.mpqueue.empty():
            msg = self.mpqueue.get()
            self.progressbar['value'] = msg[0]
            self.progressbar['maximum'] = msg[1]

            if self.label_info['text'] != msg[2]:
                # self.label_info.configure(text=msg[2])
                self.label_info.delete(0, END)
                self.label_info.insert(0, msg[2])

        if self.progressbar['value'] == self.progressbar['maximum']:
            # self.label_info.configure(text='Завершено')
            self.label_info.delete(0, END)
            self.label_info.insert(0, 'Завершено!')
            self.button_start['state'] = ACTIVE
        else:
            self.button_start['state'] = DISABLED

        self.master.after(100, self.__get_info_from_subprocess)

    def __start(self) -> None:
        self.button_start['state'] = DISABLED
        self.__save()
        self.__subprocess_close()

        if service.msword_installed():
            self.subprocess1 = Process(target=service.start_work,
                                       args=(self.mpqueue,),
                                       daemon=True)
            self.subprocess1.start()
        else:
            messagebox.showinfo(self.name_program,
                                'MS Word не установлен, файлы старого формата DOC будут пропущены без преобразования в DOCX!')
            self.subprocess1 = Process(target=service.rename_docx,
                                       args=(self.mpqueue,),
                                       daemon=True)
            self.subprocess1.start()

        self.__get_info_from_subprocess()
