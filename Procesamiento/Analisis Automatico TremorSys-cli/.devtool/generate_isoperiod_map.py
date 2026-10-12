"""
generate_isoperiod_map.py
Genera un mapa de contornos de Isoperíodos T0 (s) del proyecto Promart-CityMall
usando los datos de DataSheet.dat y guardándolo en Informe/Mapa_Isoperiodos_T0.png.
"""
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import griddata

# Leer datos de DataSheet.dat
# Formato: Lon, Lat, T0
data = []
with open(r"Informe\DataSheet.dat", "r") as f:
    for line in f:
        line = line.strip()
        if line:
            parts = [float(p) for p in line.split(",")]
            data.append(parts)

data = np.array(data)
lons = data[:, 0]
lats = data[:, 1]
t0s = data[:, 2]

# Grilla regular para interpolación
xi = np.linspace(lons.min() - 0.0001, lons.max() + 0.0001, 200)
yi = np.linspace(lats.min() - 0.0001, lats.max() + 0.0001, 200)
xi_grid, yi_grid = np.meshgrid(xi, yi)

# Interpolación cúbica / RBF
zi = griddata((lons, lats), t0s, (xi_grid, yi_grid), method='cubic')
# Relleno de bordes con linear
zi_lin = griddata((lons, lats), t0s, (xi_grid, yi_grid), method='linear')
zi = np.where(np.isnan(zi), zi_lin, zi)

plt.figure(figsize=(9, 7), dpi=300)
ax = plt.subplot(111)

# Niveles de contorno
levels = np.linspace(0.8, 1.7, 19)
cf = ax.contourf(xi_grid, yi_grid, zi, levels=levels, cmap='turbo', extend='both')
c = ax.contour(xi_grid, yi_grid, zi, levels=levels[::2], colors='black', linewidths=0.6, alpha=0.7)
ax.clabel(c, inline=True, fontsize=8, fmt='%.2f s')

# Puntos de medición
ax.scatter(lons, lats, c='white', edgecolors='black', s=55, zorder=5, linewidths=1.2)

# Etiquetas de sitios
site_labels = [f"N{i+1}" for i in range(len(lons))]
for i, txt in enumerate(site_labels):
    ax.annotate(txt, (lons[i], lats[i]), textcoords="offset points", xytext=(0, 6),
                ha='center', fontsize=8, fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="black", lw=0.6, alpha=0.85))

cbar = plt.colorbar(cf, ax=ax, orientation='vertical', pad=0.03, aspect=25)
cbar.set_label('Período Fundamental T₀ (segundos)', fontsize=10, fontweight='bold')
cbar.ax.tick_params(labelsize=8)

ax.set_title('PROYECTO PROMART - CITYMALL\nMAPA DE ISOPERÍODOS FUNDAMENTALES DEL SUELO T₀ (s)',
             fontsize=11, fontweight='bold', pad=12)
ax.set_xlabel('Longitud (°)', fontsize=9, fontweight='bold')
ax.set_ylabel('Latitud (°)', fontsize=9, fontweight='bold')
ax.tick_params(labelsize=8)
ax.ticklabel_format(useOffset=False, style='plain')
plt.grid(True, linestyle='--', alpha=0.4, color='gray')

out_path = r"Informe\Mapa_Isoperiodos_T0.png"
plt.tight_layout()
plt.savefig(out_path, dpi=300)
plt.close()
print(f"[OK] Mapa de isoperíodos generado: {out_path}")
