# -*- coding: utf-8 -*-

class SplashScreen(Window):
    """
    Заготовка для создания стартового приветственного окна
    """
    def __init__(self, parent, on_close):
        super().__init__()
        self.window = tk.Toplevel(parent)
        self.window.title("NerpaAI")
        self.window.iconbitmap(self.pic_path)
 
        self.window.resizable(False, False)
        self.window.attributes("-topmost", True)  # Убедитесь, что окно поверх всех

        # Загрузка и отображение изображения PNG
        self.image_path = get_resource_path(
            'resources\\pic\\ICO RGSH PS 400X400.png')  # Укажите путь к вашему PNG файлу
        self.image = Image.open(self.image_path)
        self.photo = ImageTk.PhotoImage(self.image)

        self.label = tk.Label(self.window, image=self.photo)
        self.label.pack(expand=True, fill=tk.BOTH)
        #self.window.update_idletasks()
        w, h = self.get_center_window(self.window)
        self.window.geometry('+{}+{}'.format(w, h))

        # Запуск таймера для закрытия окна
        self.window.after(3000, self.close)
        self.on_close = on_close

    def close(self):
        self.window.destroy()
        self.on_close()