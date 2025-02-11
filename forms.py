import tkinter as tk
from tkinter import messagebox

def obtener_horarios():
    horarios = []

    def capturar_horarios():
        nonlocal horarios
        horarios = [entrada.get() for entrada in entradas_horarios]
        root.destroy()  # Cierra la ventana una vez capturados los horarios

    def crear_entradas():
        try:
            num_horarios = int(entry_num_horarios.get())
            if num_horarios <= 0:
                raise ValueError("El número debe ser mayor que 0")
        except ValueError as e:
            messagebox.showerror("Error", f"Entrada inválida: {e}")
            return

        # Limpiar entradas anteriores
        for widget in frame_inputs.winfo_children():
            widget.destroy()

        # Crear entradas para los horarios
        entradas_horarios.clear()
        for i in range(num_horarios):
            label = tk.Label(frame_inputs, text=f"Ingrese horario {i + 1}:")
            label.grid(row=i, column=0, padx=5, pady=5)
            entry = tk.Entry(frame_inputs)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entradas_horarios.append(entry)

        # Botón para capturar horarios
        btn_guardar = tk.Button(frame_inputs, text="Guardar Horarios", command=capturar_horarios)
        btn_guardar.grid(row=num_horarios, column=0, columnspan=2, pady=10)

    # Configuración de la ventana principal
    root = tk.Tk()
    root.title("Ingreso de Horarios")

    tk.Label(root, text="Ingrese número de horarios:").pack(padx=10, pady=5)
    entry_num_horarios = tk.Entry(root)
    entry_num_horarios.pack(padx=10, pady=5)

    btn_ingresar = tk.Button(root, text="Ingresar", command=crear_entradas)
    btn_ingresar.pack(pady=10)

    frame_inputs = tk.Frame(root)
    frame_inputs.pack(padx=10, pady=10)

    entradas_horarios = []

    root.mainloop()
    return horarios
