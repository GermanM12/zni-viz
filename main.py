import os
import glob
import pandas as pd
import plotly.graph_objects as go

# 1. Montar Drive si estás en Colab, o usar cwd en local
try:
    from google.colab import drive
    drive.mount('/content/drive', force_remount=True)
    base_path = '/content/drive/MyDrive/Optimización'
except ImportError:
    base_path = os.getcwd()

# 2. Definir las dos carpetas principales
top_dirs = ['Resolución', 'Resolución&Adquisiciones']
print("Carpetas disponibles en Optimización:")
for i, d in enumerate(top_dirs):
    print(f"  {i}: {d}")
sel0 = int(input("Selecciona carpeta (número): "))
major_folder = top_dirs[sel0]

# 3. Listar escenarios dentro de esa carpeta
root_dir = os.path.join(base_path, major_folder)
scenarios = [d for d in os.listdir(root_dir)
             if os.path.isdir(os.path.join(root_dir, d))]
print(f"\nEscenarios en '{major_folder}':")
for i, s in enumerate(scenarios):
    print(f"  {i}: {s}")
sel1 = int(input("Selecciona escenario (número): "))
scenario = scenarios[sel1]

# 4. Listar .xlsx en el escenario
scenario_folder = os.path.join(root_dir, scenario)
files = glob.glob(os.path.join(scenario_folder, '*.xlsx'))
print(f"\nArchivos en '{scenario}':")
for i, f in enumerate(files):
    print(f"  {i}: {os.path.basename(f)}")
sel2 = int(input("Selecciona archivo (número): "))
excel_path = files[sel2]

# 5. Cargar y limpiar los datos (filas 2–1302)
df = pd.read_excel(excel_path, header=0)
df_data = df.iloc[1:1302].copy()

# 6. Mapear columnas por posición
cols         = df.columns
dept_col     = cols[5]   # F: Departamento
users_col    = cols[8]   # I: No. de usuarios
lcoe_col     = cols[33]  # AH: LCOE (mCOP/kWh)
gen_col      = cols[29]  # AD: Generación (kWh)
subsidio_col = cols[20]  # U: Subsidio (mCOP)
solar_col    = cols[23]  # X: Fracción solar
biomasa_col  = cols[24]  # Y: Fracción biomasa
diesel_col   = cols[25]  # Z: Fracción diésel
subratio     = 0.5       # 1 - SubRatio

# 7. Convertir a numérico y filtrar
for c in [users_col, lcoe_col, gen_col, subsidio_col, solar_col, biomasa_col, diesel_col]:
    df_data[c] = pd.to_numeric(df_data[c], errors='coerce')
df_data = df_data[
    df_data[dept_col].astype(str).str.strip().astype(bool) &
    (df_data[users_col] > 0)
]

# 8. Agregar métricas y calcular proporciones
def agg(x):
    total_users = x[users_col].sum()
    lcoe = ((x[lcoe_col]*1000)*x[users_col]).sum() / total_users
    gen_pu = x[gen_col].sum() / total_users
    pago = ((x[gen_col]*(x[lcoe_col]*1000)*(1-subratio)).sum()) / total_users
    subs = (x[subsidio_col]*1000).sum()
    sumX, sumY, sumZ = x[solar_col].sum(), x[biomasa_col].sum(), x[diesel_col].sum()
    tot = sumX + sumY + sumZ
    return pd.Series({
        'LCOE_COP_kWh': lcoe,
        'Generacion_per_user': gen_pu,
        'Pago_prom_user_COP': pago,
        'Subsidio_total_COP': subs,
        'Solar_prop':   sumX/tot if tot>0 else 0,
        'Biomasa_prop': sumY/tot if tot>0 else 0,
        'Diesel_prop':  sumZ/tot if tot>0 else 0
    })

df_agg = df_data.groupby(dept_col).apply(agg).reset_index()

# 9. Gráfico 1: métricas interactivo
metrics = ['LCOE_COP_kWh','Generacion_per_user','Pago_prom_user_COP','Subsidio_total_COP']
fig1 = go.Figure()
for i, m in enumerate(metrics):
    fig1.add_bar(x=df_agg[dept_col], y=df_agg[m], name=m, visible=(i==0))
buttons = [
    dict(label=m, method='update',
         args=[{'visible':[j==i for j in range(len(metrics))]},
               {'title':f'{m} por Departamento'}])
    for i,m in enumerate(metrics)
]
fig1.update_layout(
    updatemenus=[dict(active=0, buttons=buttons, x=0, y=1.2)],
    title=metrics[0]+' por Departamento',
    xaxis_title='Departamento', yaxis_title='Valor',
    margin=dict(t=120,b=150)
)

# 10. Gráfico 2: proporciones apiladas
fig2 = go.Figure()
fig2.add_bar(x=df_agg[dept_col], y=df_agg['Solar_prop'],   name='Solar',   marker_color='blue')
fig2.add_bar(x=df_agg[dept_col], y=df_agg['Biomasa_prop'], name='Biomasa', marker_color='green')
fig2.add_bar(x=df_agg[dept_col], y=df_agg['Diesel_prop'],  name='Diésel',  marker_color='red')
fig2.update_layout(
    barmode='stack',
    title='Proporción de fuentes por Departamento',
    xaxis_title='Departamento', yaxis_title='Proporción',
    margin=dict(t=120,b=150)
)

# 11. Mostrar ambos
fig1.show()
fig2.show()
