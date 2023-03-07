# Приложение Киоск Инфостенд
### Запуск
Приложение запускается Планировщиком задач при входе пользователя I-PRO-INFO-TV в систему.
Приложение запускается с параметрами, пример запуска:
```
kiosk.exe “Путь к файлу.pptx” 2 5 300
```
```angular2html
где:
2 – Номер слайда меню (по умолчанию – 2)
5 - частота смены слайдов (по умолчанию - 5 сек)
300 - время до перехода в режим бездействия (автоматического переключения слайдов) - (по умолчанию - 300 сек)

```
Обязательный параметр «Путь к файлу», остальные по умолчанию.

### Исполнение
Приложение запускает браузер Chrome в режиме киоска и презентацию в полноэкранном автоматическом режиме (слайд-шоу). 
Приложение проверяет наличие обновленной версии презентации "Путь к файлу.pptx". При обнаружении файла с датой, отличной от даты текущего файла презентации, приложение копирует этот файл в свою папку и перезапускает презентацию.
При перемещении курсора мыши презентация переходит на слайд меню (№2) в ручной режим. В этом режиме управление презентацией выполняется мышью или клавиатурой. 
Автоматический режим включается через установленное время до перехода в режим бездействия мыши или при нажатии клавиши Esc. При переходе по гиперссылке из презентации целевая страница открывается в браузере Chrome.
При попытке закрыть презентацию или браузер они открываются автоматически.
### Выход
Выход из программы комбинацией клавиш ```[Ctrl Alt Shift -]```
