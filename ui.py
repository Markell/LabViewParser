import shutil
import os.path
from os import system
from rich.table import Table
from rich.console import Console
from datetime import date
from pathlib import Path

version = '2.0'

console = Console()
cwd = os.getcwd()

xls_file = os.path.join(cwd, "template.xlsx")
txt_file = os.path.join(cwd, "output.txt")
frequency_path = os.path.join(cwd, "data.txt")

work_dir = ""
date = date.today().strftime("%d.%m.%Y")


def create_environment(frequency_range):
    global work_dir
    directory = os.path.join(cwd, ("Cбор " + date + " " + frequency_range))
    Path(directory).mkdir(parents=True, exist_ok=True)
    work_dir = directory


class UInterface:

    def generate_table(self, array, position):
        system('cls')

        if position == 0:
            position = "[green]Данные для ближней катушки[/green]"
        else:
            position = "[yellow]Данные для дальней катушки[/yellow]"

        table = Table(title=position)
        table.add_column("Позиция")
        table.add_column("Cos  2 чужая")
        table.add_column("Sin  2 чужая")
        table.add_column("Фаза 2 чужая")
        table.add_column("Cos  2 своя")
        table.add_column("Sin  2 своя")
        table.add_column("Фаза 2 своя")
        table.add_column("Cos  6 чужая")
        table.add_column("Sin  6 чужая")
        table.add_column("Фаза 6 чужая")
        table.add_column("Cos  6 своя")
        table.add_column("Sin  6 своя")
        table.add_column("Фаза 6 своя")

        for row in range(len(array)):
            table.add_row(f"{row + 1}",
                          f"{array[row][0]}",
                          f"{array[row][1]}",
                          f"{array[row][2]}",
                          f"{array[row][3]}",
                          f"{array[row][4]}",
                          f"{array[row][5]}",
                          f"{array[row][6]}",
                          f"{array[row][7]}",
                          f"{array[row][8]}",
                          f"{array[row][9]}",
                          f"{array[row][10]}",
                          f"{array[row][11]}"
                          )
        console.print(table)

    def welcome_output(self):
        console.print("\n",
                      "                 ░██████████████████████████████▒           \n",
                      "                                                            \n",
                      "                 ▓▓▓▓▓▓▓▓▒   ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▒     \n",
                      "                                                            \n",
                      "                 ██████████████████▒  ▓██████████▓  ▒████   \n",
                      "                                                            \n",
                      "                 ██████  ░▓███▓              ▓█████████████▒\n",
                      "                                                            \n",
                      "                 ▒▒▒▒▒▒▒▒▒▒▒▒▒▒               ▒▒▒▒▒▒▒▒▒▒▒▒▒ \n",
                      "                 ▓▓▓▓▓▓▓▓▓▓▓▓▓▒              ▒▓▓▓▓▓▓▓▓▓▓▓▓▓ \n",
                      "                                                            \n",
                      "                 ████   █████████   █████████████████████   \n",
                      "                                                            \n",
                      "                 ▓▓▓▓▓▓▓▓▓▓▓   ▓▓▓▓▓▓▓▓▓▓▓▒  ▒▓▓▓▓▓▓▓▓▒     \n",
                      "                                                            \n",
                      "                 ▒▒▒▒   ▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒             \n",
                      "                                                            \n",
                      "                 █████████████▓                             \n",
                      "                                                            \n",
                      "                 ▓▓▓▓▓▓▒  ▒▓▓▓▒                             \n",
                      "                 ▓▓▓▓▓▓▒  ▒▓▓▓▒                             \n",
                      "                                                            \n",
                      "                 █████████████▓                           \n\n",
                      "                         LabViewParser v", version
                      )
        console.print("Вас приветствует программа парсинга LabView [u]MeDeCoPan[/u] для двухчастотного анализа.\n")

        frequency_file = Path(frequency_path)

        if frequency_file.is_file():
            file = open(frequency_path, "r")
            file_data = file.read()
            file.close()
            print(" В прошлый раз вы использовали пару частот ", file_data,
                  " использовать данные повторно?\n Нажмите Enter чтобы продожить со старой парой частот",
                  "\n Введите любой символ чтобы выбрать новую пару частот")
            user_answer = input(">")
            if user_answer:
                file = open(frequency_path, "w")
                self.get_frequency(file)
            else:
                create_environment(file_data)
        else:
            file = open(frequency_path, "w")
            self.get_frequency(file)

    def get_frequency(self, file):
        while True:
            console.print("\nВведите порядковый номер используемой пары частот из списка ниже\n",
                          "\n1)  8117/7837", "\n2)  8333/8013", "\n3)  8333/8065", "\n4)  8418/8117",
                          "\n5)  8621/8333", "\n6)  8681/8333", "\n7)  8741/8418", "\n8)  8929/8621",
                          "\n9)  9091/8741", "\n10) 9259/8961", "\n11) 9470/9091", "\n12) 9579/9259",
                          "\n13) 9921/9579", "\n14) 10000/9615", "\n15) 10288/9921")

            user_input = console.input("\n>")
            two_frequency_list = {
                '1': [8117, 7837], '2': [8333, 8013], '3': [8333, 8065], '4': [8418, 8117], '5': [8621, 8333],
                '6': [8681, 8333], '7': [8741, 8418], '8': [8929, 8621], '9': [9091, 8741], '10': [9259, 8961],
                '11': [9470, 9091], '12': [9579, 9259], '13': [9921, 9579], '14': [10000, 9615], '15': [10288, 9921]
            }
            try:
                create_environment(str(two_frequency_list[user_input]))
                break
            except KeyError as e:
                console.print("\n[red]Не правильно указан номер[/red]")
                continue

        file.write(str(two_frequency_list[user_input]))
        file.close()

    def path_check(self, path):
        file = Path(path)
        if file.is_file():
            return True
        console.print("[red]Файл не найден![/red]")
        return False

    def txt_path_input(self):
        global txt_file
        if self.path_check(txt_file):
            console.print("[green]output.txt обнаружен![/green]")
            return txt_file
        else:
            while True:
                txt_file = console.input("Введите путь до файла [i]output.txt[/i] выбранный в программе LabVIEW"
                                         " или нажмите [u]Enter[/u] (если файл находится в директории программы)\n>")

                if not txt_file:
                    txt_file = os.path.join(cwd, "output.txt")
                elif not ("output.txt" in txt_file):
                    txt_file = os.path.join(txt_file, "output.txt")

                if self.path_check(txt_file):
                    return txt_file
                else:
                    continue

    def xls_path_input(self):
        global xls_file

        if self.path_check(xls_file):
            console.print("[green]template.xlsx обнаружен![/green]")
        else:
            while True:
                xls_file = console.input("Введите путь до файла [i]template.xlsx[/i] он будет выбран в качестве шаблона"
                                         " или нажмите [u]Enter[/u] (если файл находится в директории программы)\n>")

                if not xls_file:
                    xls_file = os.path.join(cwd, "template.xlsx")
                elif not ("output.xlsx" in xls_file):
                    xls_file = os.path.join(xls_file, "template.xlsx")

                if self.path_check(xls_file):
                    break
                else:
                    continue

    def tor_number_input(self):
        global xls_file, work_dir
        while True:
            tor_number = console.input("Ведите порядковый номер тора\n>")
            if tor_number: break
            console.print("[red]Поле не должно быть пустым![/red]\n")

        name = "тор М" + str(tor_number) + " (" + date + ")" + ".xlsx"

        if not Path(work_dir + "\\" + name).is_file():
            shutil.copy(xls_file, work_dir)
            os.rename((work_dir + "\\template.xlsx"), work_dir + "\\" + name)
        else:
            console.print("[red]Файл с таким именем уже существет!\n[/red]")
            while True:
                console.print("[yellow]Создать новый файл            (введите 1)\n"
                              "Внести измения в существующий (введите 2)[/yellow]")
                i = console.input(">")
                if i == "1":
                    os.remove(work_dir + "\\" + name)
                    shutil.copy(xls_file, work_dir)
                    os.rename((work_dir + "\\template.xlsx"), work_dir + "\\" + name)
                    break
                elif i == "2":
                    break
                else:
                    console.print("[red]Введите корректный номер операции[/red]")
                    continue

        console.input("[green]Все готово[/green]\n"
                      "[yellow]Убедитесь что новые данные записаны в output.txt[/yellow]\n"
                      "В противном случае старые данные будут записаны в новый сбор или программа закроется\n"
                      "Для продолжения нажмите [u]Enter[/u]")

        return os.path.join(work_dir, name)
