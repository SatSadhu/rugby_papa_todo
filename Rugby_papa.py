import sys
import tkinter as tk
from tkinter import ttk
import pickle
from tkinter import simpledialog, messagebox
from pathlib import Path
import os
# pyright: reportMissingModuleSource=false
import xlsxwriter

# Crear ventana principal
root = tk.Tk()
root.title("Tabla de estadísticas de rugby")
root.geometry("1366x768")
root.resizable(False, False)

#Texto 1er tiempo
label = tk.Label(text="1er Tiempo")
label.place(y=0, x=700)

# Crear un widget Treeview
treeview = ttk.Treeview(root)
treeview.place(y=20, x=0, width=1366, height=2000)

# Configurar filas y columnas para que se ajusten al tamaño de la ventana
root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

# Insertar columnas y filas de ejemplo
treeview["columns"] = ("Tiempo", "Tackles", "Arriba", "Abajo", "H.Interno", "H.Externo", "Adelante LV", "Misma LV",
                       "Atras LV", "Positivo", "Neutro", "Negativo", "Doble Tackle", "Errados")

names_list = []

for column in treeview["columns"]:
    treeview.column(column, width=60)

treeview.column("#0", width=160)
treeview.column("Tiempo", width=110)

treeview.heading("#0", text="Jugador")
treeview.heading("Tiempo", text="Tiempo")
treeview.heading("Tackles", text="Tackles")
treeview.heading("Arriba", text="Arriba")
treeview.heading("Abajo", text="Abajo")
treeview.heading("H.Interno", text="H.Interno")
treeview.heading("H.Externo", text="H.Externo")
treeview.heading("Adelante LV", text="Adelante LV")
treeview.heading("Misma LV", text="Misma Linea")
treeview.heading("Atras LV", text="Linea Atras LV")
treeview.heading("Positivo", text="Positivo")
treeview.heading("Neutro", text="Neutro")
treeview.heading("Negativo", text="Negativo")
treeview.heading("Doble Tackle", text="Doble Tackle")
treeview.heading("Errados", text="Errados")

# Definir una variable booleana para llevar registro de si la fila ha sido seleccionada
fila_seleccionada = {}


def opciones():
    menu = tk.Menu(root)
    root.config(menu=menu)
    opciones = tk.Menu(menu, tearoff=0)
    opciones.add_command(label="Guardar", command=guardar)
    opciones.add_command(label="Agregar Jugadores", command=agregar_jugadores)
    opciones.add_command(label="Agregar Tiempo", command=agregar_tiempo)
    opciones.add_command(label="Ir a 2do tiempo", command=nuevo_2do_tiempo)
    opciones.add_command(label="Sacar Totales", command=total)
    menu.add_cascade(label="Opciones", menu=opciones)


# Función que se ejecuta al hacer clic con el botón izquierdo
def on_left_click(event):
    opciones()
    global previous_value, fila_seleccionada

    # Obtener la fila y columna del click
    row = treeview.identify_row(event.y)
    column = treeview.identify_column(event.x)

    # Obtener índice de la columna
    column_index = int(str(column).replace("#", "")) - 1

    # Obtener valor actual y cambiarlo
    valor_actual = treeview.set(row, column_index)

    # Obtener el número de la columna y el valor actual
    col_num = int(str(column).replace("#", ""))
    value = treeview.item(row)["values"][col_num - 1]

    # Verificar si la fila ha sido seleccionada antes
    if row in fila_seleccionada and fila_seleccionada[row]:
        if column == "#1":
            pass
        else:
            # Sumar 1 al valor actual
            try:
                value += 1
            except TypeError:
                pass
    else:
        # Si es el primer click en la fila, marcarla como seleccionada sin cambiar el valor
        fila_seleccionada[row] = True

    # Actualizar el valor en el Treeview
    treeview.set(row, column, value)

    # Guardar el valor actual como el valor anterior
    previous_value = value


def on_right_click(event):
    global previous_value

    # Obtener la fila y columna del click
    row = treeview.identify_row(event.y)
    column = treeview.identify_column(event.x)

    # Obtener el número de la columna y el valor actual
    col_num = int(str(column).replace("#", ""))
    value = treeview.item(row)["values"][col_num - 1]

    # Restar 1 al valor actual
    if column == "#1":
        pass
    else:
        # Sumar 1 al valor actual
        try:
            value -= 1
        except TypeError:
            pass

    # Actualizar el valor en el Treeview
    treeview.set(row, column, value)

    # Guardar el valor actual como el valor anterior
    previous_value = value


# Vincular la función `column_click` al evento "<Button-1>" y "<Button-3>" en el widget Treeview
treeview.bind("<Button-1>", on_left_click)
treeview.bind("<Button-3>", on_right_click)


def guardar():
    # Obtener todos los valores de las filas y columnas
    values = []
    for item in treeview.get_children():
        values.append(treeview.item(item)["values"])

    # Obtener los nombres de los jugadores
    player_names = []
    for item in treeview.get_children():
        player_names.append(treeview.item(item)["text"])

    # Crear el archivo xlsx y hoja de trabajo
    if Path('Rugby_Excel_1erT.xlsx').is_file():
        workbook = xlsxwriter.Workbook("Rugby_Excel_2doT.xlsx")
        worksheet = workbook.add_worksheet()
    else:
        workbook = xlsxwriter.Workbook("Rugby_Excel_1erT.xlsx")
        worksheet = workbook.add_worksheet()

    # Escribir el nombre de la columna "JUGADORES" en la celda A1
    worksheet.write(0, 0, "JUGADORES")

    # Escribir los nombres de los jugadores en la primera columna
    for i, player_name in enumerate(player_names):
        worksheet.write(i+1, 0, player_name)

    # Escribir los encabezados de las columnas en la primera fila
    headers = ["Tiempo", "Tackles", "Arriba", "Abajo", "H.Interno", "H.Externo", "Adelante", "Misma LV", "Atras LV",
               "Positivo LV", "Neutro LV", "Negativo LV", "Doble Tackle", "Errados"]
    for i, header in enumerate(headers):
        worksheet.write(0, i+1, header)

    # Escribir los valores de las filas y columnas en el archivo xlsx
    for i, row in enumerate(values):
        for j, value in enumerate(row):
            worksheet.write(i+1, j+1, str(value))

    # Cerrar el archivo xlsx
    workbook.close()

    # Confirmar que se guardó el archivo correctamente
    messagebox.showinfo("Guardado",
                        "Los datos se han guardado correctamente en\n(Rugby_Excel_1erT.xlsx / Rugby_Excel_2doT.xlsx)")


def agregar_jugadores():
    global entry_dato
    global nueva
    nueva = tk.Toplevel(root)
    nueva.title("Agregar Jugadores")
    nueva.resizable(False, False)
    nueva.geometry("1366x768")

    label = tk.Label(nueva, text="Cantidad de Jugadores")
    label.grid(row=0, column=0)

    entry_dato = tk.StringVar()
    entry = ttk.Entry(nueva, textvariable=entry_dato)
    entry.grid(row=1, column=0)

    boton = ttk.Button(nueva, text="Agregar", command=lambda: entrys(nueva, names))
    boton.grid(row=2, column=0)


names = []


# Función para crear los entrys
def entrys(nueva, names):
    # Crear 5 entrys dentro de un bucle
    for i in range(int(entry_dato.get())):
        # Crear un label con el número del entry
        label = tk.Label(nueva, text=f"Jugador {i + 1}")
        label.grid(row=i, column=3)

        # Crear el entry y guardarlo en la lista de valores
        entry = tk.Entry(nueva)
        entry.grid(row=i, column=4)
        names.append(entry)

    boton = tk.Button(nueva, text="Agregar Jugadores", command=lambda: mostrar_valores(names))
    boton.grid(row=25, column=3, columnspan=2)


# Función para mostrar los valores de los entrys
def mostrar_valores(names):
    # Recorrer la lista de valores y mostrarlos en la consola
    for i, entry in enumerate(names):
        names_list.append(entry.get())
    #Agregar jugadores
    for name in names_list:
        treeview.insert("", 0, text=name, values=(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
    nueva.destroy()


def agregar_tiempo():
    # Obtener el jugador seleccionado
    item = treeview.focus()
    jugador = treeview.item(item)["text"]

    # Obtener el tiempo a agregar
    tiempo = simpledialog.askstring("Ingresar Tiempo", "Ingrese Tiempo (HH:MM:SS):")

    # Obtener la fila del jugador seleccionado y establecer el tiempo
    for row in treeview.get_children():
        if treeview.item(row)["text"] == jugador:
            # Obtener el tiempo existente
            tiempo_existente = treeview.item(row)["values"][0]

            # Concatenar el tiempo existente con el nuevo tiempo ingresado
            if tiempo == "":
                pass
            else:
                if tiempo_existente == 0:
                    nuevo_tiempo = f"{tiempo}"
                    if nuevo_tiempo == None:
                        nuevo_tiempo.replace("None", "")
                else:
                    nuevo_tiempo = f"{tiempo_existente} / {tiempo}"

                # Actualizar el valor en la tabla
                treeview.set(row, "Tiempo", nuevo_tiempo)
                break


def guardar_datos():
    # Obtener todos los valores de las filas y columnas
    values = []
    for item in treeview.get_children():
        values.append(treeview.item(item)["values"])

    mensaje = messagebox.askyesno("Guardar", "Los datos se estan por perder, esta seguro que guardo correctamente?",
                                  default="no")
    if mensaje:
        if Path('datos.bin').is_file():
            os.remove("datos.bin")
        sys.exit()
    else:
        pass

    # Serializar los datos y guardarlos en un archivo binario
    with open("datos.bin", "wb") as f:
        pickle.dump(values, f)


def cargar_datos():
    try:
        # Cargar los datos desde el archivo binario
        with open("datos.bin", "rb") as f:
            values = pickle.load(f)

        # Insertar las filas y columnas en el Treeview
        for row in values:
            treeview.insert("", tk.END, values=row)

    except FileNotFoundError:
        pass


# Vincular la función `guardar_datos` al evento "<Destroy>" de la ventana principal
root.bind("<Destroy>", lambda event: guardar_datos())

# Cargar los datos cuando la ventana se abra
cargar_datos()


def total():
    # Obtener todas las filas
    rows = treeview.get_children()
    columns = treeview["columns"]

    # Buscar la fila que contiene la cadena "total" en la columna "Jugadores #0"
    for row in rows:
        if treeview.item(row, "text") == "TOTAL":
            # Eliminar la fila si existe
            treeview.delete(row)

    # Diccionario para almacenar los totales de cada columna
    totals = {}

    # Iterar a través de todas las filas
    for child in treeview.get_children():
        # Obtener valores de la fila actual
        values = treeview.item(child, 'values')

        # Iterar a través de cada columna
        for i, val in enumerate(values):
            # Si la columna no está en el diccionario, inicializar el valor en cero
            if i not in totals:
                totals[i] = 0

            # Sumar el valor actual al total de la columna
            if val:
                if i == 0:
                    pass
                else:
                    totals[i] += int(val)

    # Agregar la fila de totales
    total_values = [totals.get(i, '') for i in range(len(treeview["columns"]))]
    treeview.insert("", tk.END, text="TOTAL", values=total_values)


def nuevo_2do_tiempo():
    msg = messagebox.askyesno("Cambiar a 2do Tiempo", "Esta seguro que desea cambiar al 2do Tiempo?"
                                                      "\nTodos los datos del 1er Tiempo seran borrados")
    if msg:
        treeview.delete(*treeview.get_children())
        # Texto 2do tiempo
        label = tk.Label(text="2do Tiempo")
        label.place(y=(-1), x=700)
        names_list.clear()
        names.clear()

root.mainloop()
