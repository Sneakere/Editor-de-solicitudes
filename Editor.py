import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import os 
# Variables globales

global archivo_cargado, repofiltrado, estilos_filtrados, repo, n_traspaso, n_stock, n_rotacion, n_stock_bodega, filtro_marca


# Ventana raiz 
root = tk.Tk()
root.geometry("750x700")
root.config(background="#edebe6")
root.title("Editor de solicitudes")
root.resizable(width=False, height=False)

def reset_app():
    global archivo_cargado, repofiltrado, estilos_filtrados, repo
    archivo_cargado = None
    repofiltrado = pd.DataFrame()
    estilos_filtrados = []
    repo = pd.DataFrame()
    
    etiqueta_archivo.config(text="Archivo seleccionado: ")
    texto_resultado.config(state="normal")
    texto_resultado.delete(1.0, tk.END)
    texto_resultado.config(state="disabled")

def editarexcel():
    global archivo_cargado
    archivo = filedialog.askopenfilename(title="Seleccionar archivo")
    if archivo:
        archivo_cargado = archivo
        nombre_archivo = archivo.split("/")[-1]  
        etiqueta_archivo.config(text=f"Archivo seleccionado: {nombre_archivo}")
def filtrar_archivo():
        global archivo_cargado
        if archivo_cargado:
            try:
                hoja1 = pd.read_excel(archivo_cargado, sheet_name="Hoja1")
                hoja1.columns = hoja1.columns.str.rstrip()
                hoja1["ESTILOCOLOR"] = hoja1["ESTILO"].astype(str) + hoja1["COLOR"].astype(str)
                hoja1 = hoja1.applymap(lambda x:x.strip() if isinstance(x, str) else x)
                #hoja1["ROTACION"] = hoja1["ROTACION"].astype(str).apply(lambda x: x.rstrip())                
                hoja1["ROTACION"].replace(["NA", "Inf", "nan", "inf", np.inf, np.nan], 0, inplace=True)

                global repo

                repo = pd.read_excel(archivo_cargado, sheet_name="REPO")
                repo["ESTILO"] = repo["SKU"].astype(str).str.slice(0, 6).astype("Int64")
                repo["ESTILOCOLOR"] = repo["ESTILO"].astype(str) + repo["COLOR"].astype(str)
                
                global n_traspaso
                global n_stock
                global n_rotacion
                global n_stock_bodega
                global filtro_marca
                n_traspaso = int(entry_traspaso.get())
                n_stock = int(entry_stock.get())
                n_rotacion = int(entry_rotacion.get())
                n_stock_bodega = int(entry_stock_bodega.get())
                filtro_marca = str(entry_filtro_marca.get())
                



                # ELIMINAR FILAS QUE TENGAN SKUS EN TRASPASOS
                estilo_color_entraspaso = hoja1.loc[hoja1["EN TRASPASOS"].astype(float) > n_traspaso, "ESTILOCOLOR"].tolist()
                # ELIMINAR FILAS QUE TENGAN SKUS CON STOCK EN TIENDA SUPERIOR O IGUAL A n
                estilo_color_stock_tienda = hoja1.loc[hoja1["STOCK TIENDA"].astype(float) >= n_stock, "ESTILOCOLOR"].tolist()
                # ELIMINAR FILAS QUE TENGAN SKUS CON ROTACION SUPERIOR A n
                estilo_color_rotacion = hoja1.loc[hoja1["ROTACION"].astype(float) >= n_rotacion, "ESTILOCOLOR"].tolist()
                # ELIMINAR SKUS CUYO STOCK EN BODEGA SEA MENOR A n
                estilo_color_stock_bodega = hoja1.loc[hoja1["CURVAS_DISPONIBLES_ESTILO"].astype(float) <= n_stock_bodega, "ESTILOCOLOR"].tolist()
                # ELIMINAR MARCA SELECCIONADA
                filtro_marca_id = hoja1.loc[hoja1["MARCA_ID"] == filtro_marca, "MARCA_ID"].unique().tolist()
                # TODOS LOS ESTILOS A EXCLUIR

                global estilos_filtrados

                estilos_filtrados = list(set(estilo_color_entraspaso + estilo_color_stock_tienda + 
                    estilo_color_rotacion + estilo_color_stock_bodega + filtro_marca_id))
                estilos_filtrados = list(estilos_filtrados)
                print(estilos_filtrados)
                global repofiltrado

                repofiltrado = hoja1[hoja1["ESTILOCOLOR"].isin(estilos_filtrados) | hoja1["MARCA_ID"].isin(estilos_filtrados)]
                columnas_a_cargar = ["ESTILO", "MARCA_ID", "COLOR", "ROTACION", "STOCK TIENDA", "CURVAS_DISPONIBLES_ESTILO", "EN TRASPASOS"]
                repofiltrado = repofiltrado[columnas_a_cargar]

                #print(repofiltrado)


                # Mostrar los estilos filtrados
                texto_resultado.config(state="normal")  
                texto_resultado.delete(1.0, tk.END)  
                texto_resultado.insert(tk.END, repofiltrado.to_string(index=False))  
                texto_resultado.config(state="disabled")  


            except Exception as e:
                resultado.set(f"Error al cargar el archivo: {str(e)}")


def preparar_csv():
    global archivo_cargado, repofiltrado, estilos_filtrados, repo
    if not repofiltrado.empty:
        repofiltrado = repo[~repo["ESTILOCOLOR"].isin(estilos_filtrados)]
        columnas_a_cargar = ["SKU", "TIENDA_ORIGEN", "TIENDA_DESTINO", "QTY"]
        repofiltrado = repofiltrado[columnas_a_cargar]
        global nombre_archivo
        nombre_archivo = simpledialog.askstring("Guardar CSV", "Ingresa el nombre del archivo CSV:")
        if nombre_archivo:
            nombre_archivo = str(nombre_archivo) + ".csv"
            repofiltrado.to_csv(nombre_archivo, index=False, header=None)

        # Abre un cuadro de diálogo para seleccionar la carpeta de destino
        folder_selected = filedialog.askdirectory(title="Seleccionar carpeta para guardar CSV")

        if folder_selected:
            # Comprueba si el nombre del archivo ya contiene la extensión .csv
            if not nombre_archivo.endswith('.csv'):
                nombre_archivo += ".csv"

            # Combina la carpeta seleccionada y el nombre del archivo
            ruta_completa = os.path.join(folder_selected, nombre_archivo)

            repofiltrado.to_csv(ruta_completa, index=False, header=None)


            # Resetear entries

            archivo_cargado = None

            entry_traspaso.delete(0, tk.END)
            entry_traspaso.insert(0, valor_traspaso_por_defecto)

            entry_stock.delete(0, tk.END)
            entry_stock.insert(0, valor_stock_por_defecto)

            entry_rotacion.delete(0, tk.END)
            entry_rotacion.insert(0, valor_rotacion_por_defecto)

            entry_stock_bodega.delete(0, tk.END)
            entry_stock_bodega.insert(0, valor_stock_bodega_por_defecto)

            entry_filtro_marca.delete(0, tk.END)
            entry_filtro_marca.insert(0, "")

            reset_app()
            print("Guardado!")

    else:
        pass



# Titulo
frame_titulo = tk.Frame(root, background="#edebe6")
frame_titulo.pack()
titulo = tk.Label(frame_titulo, text="Editor de solicitudes")
titulo.config(background="#edebe6", font=('Helvetica bold', 15))
titulo.pack(pady=3)

# Cargar archivo .xlsx para editar
frame_cargarexcel = tk.Frame(root, background="#edebe6")
frame_cargarexcel.pack()
boton_abrirexcel = tk.Button(frame_cargarexcel, text="Cargar Excel", command=editarexcel, background="#edebe6")
boton_abrirexcel.config(background="white")
boton_abrirexcel.pack()

# Nombre del archivo seleccionado.
etiqueta_archivo = tk.Label(root, text="Archivo seleccionado: ", background="#edebe6")
etiqueta_archivo.pack(pady=5)

# Frame descripcion de editor
frame_descripcion = tk.Frame(root, background="#edebe6")
frame_descripcion.pack()

# Crear el marco frame_variables
frame_variables = tk.Frame(root)
frame_variables.config(background="#edebe6")
frame_variables.pack()

# Crear frame mensaje_error
frame_error = tk.Frame(root)
frame_error.config(background="#edebe6")
frame_error.pack()

# Crear el frame boton_editar_excel
frame_boton_editar_excel = tk.Frame(root)
frame_boton_editar_excel.config(background="#edebe6")
frame_boton_editar_excel.pack()

# Crear frame visualizador_excel
visualizador_excel = tk.Frame(root)
visualizador_excel.config(background="#edebe6")
visualizador_excel.pack()

# Descripcion 
label_descripcion = tk.Label(frame_descripcion, text="Ingresa las variables consideradas para desechar estilos solicitados")
label_descripcion.config(background="#edebe6")
label_descripcion.pack()

# Valores por defecto
valor_traspaso_por_defecto = 99999
valor_rotacion_por_defecto = 99999
valor_stock_por_defecto = 99999
valor_stock_bodega_por_defecto = 0

# Crear una etiqueta para n_traspaso
label_traspaso = tk.Label(frame_variables, text="EN TRASPASOS (>):")
label_traspaso.config(background="#edebe6")
label_traspaso.grid(row=2, column=0, sticky="w")
entry_traspaso = tk.Entry(frame_variables, width=7)
entry_traspaso.insert(0, valor_traspaso_por_defecto)
entry_traspaso.grid(row=2, column=1, sticky="e", padx=10)

# Crear una etiqueta para n_stock
label_stock = tk.Label(frame_variables, text="STOCK EN  TIENDA (≥):")
label_stock.config(background="#edebe6")
label_stock.grid(row=2, column=2, sticky="w")
entry_stock = tk.Entry(frame_variables, width=7)
entry_stock.insert(0, valor_stock_por_defecto)
entry_stock.grid(row=2, column=3, sticky="e", padx=10)

# Crear una etiqueta para n_rotacion
label_rotacion = tk.Label(frame_variables, text="ROTACION (≥):")
label_rotacion.config(background="#edebe6")
label_rotacion.grid(row=3, column=0, sticky="w")
entry_rotacion = tk.Entry(frame_variables, width=7)
entry_rotacion.insert(0, valor_rotacion_por_defecto)
entry_rotacion.grid(row=3, column=1, sticky="e", padx=10)

# Crear una etiqueta para n_stock_bodega
label_stock_bodega = tk.Label(frame_variables, text="CURVAS DISPONIBLES (≤):")
label_stock_bodega.config(background="#edebe6")
label_stock_bodega.grid(row=3, column=2, sticky="w")
entry_stock_bodega = tk.Entry(frame_variables, width=7)
entry_stock_bodega.insert(0, valor_stock_bodega_por_defecto)
entry_stock_bodega.grid(row=3, column=3, sticky="e", padx=10)

# CREAR UNA ETIQUETA PARA filtro_marca
label_filtro_marca = tk.Label(frame_variables, text="MARCA_ID:")
label_filtro_marca.config(background="#edebe6")
label_filtro_marca.grid(row=4, column=0, sticky="w")
entry_filtro_marca = tk.Entry(frame_variables, width=7)
entry_filtro_marca.grid(row=4, column=1, sticky="e", padx=10)

# Crear un boton para editar excel
boton_editar = tk.Button(frame_boton_editar_excel, text="Modificar", command=filtrar_archivo, background="white")
boton_editar.pack(pady=5)

# Mensaje error al modificar
resultado = tk.StringVar()
resultado_label = tk.Label(frame_error, textvariable=resultado)
resultado_label.config(fg="red", background="#edebe6")
resultado_label.pack()

descripcion_visualizador = tk.Label(visualizador_excel, text="Los estilos eliminados serán: ", background="#edebe6")
descripcion_visualizador.pack(fill="both", expand=True)

texto_resultado = tk.Text(visualizador_excel, width="88", height="21")
texto_resultado.pack()

boton_obtener_archivo = tk.Button(visualizador_excel, text="Obtener archivo", command=preparar_csv, background="white")
boton_obtener_archivo.pack(pady=5)

root.mainloop()