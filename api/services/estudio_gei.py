#estudio_gei.py
import os
import sys

class EstudioGEI:
    def __init__(self, hds_data, unidad_base="kg"):
        """
        hds_data: lista de diccionarios con info de las sustancias (del excel aplanado)
        unidad_base: unidad base a la que convertiremos todos los consumos (por defecto kg)
        """
        self.hds_data = hds_data
        self.unidad_base = unidad_base
       

        # Crear un índice rápido por nombre de componente para encontrar PCG
        self.gei_info = self._extraer_gei_info()

    def _extraer_gei_info(self):
        """
        Crea un diccionario {nombre_componente: PCG} solo para las sustancias GEI.
        """
        gei_dict = {}
        # Asumimos que "Sujeta a GEI" es True si la sustancia es GEI.
        # Además, las filas del excel aplanado pueden tener múltiples filas por sustancia (por cada componente).
        # Debemos filtrar las filas que tengan "Sujeta a GEI" = True y que tengan "Potencial de Calentamiento Global" no None.
        for row in self.hds_data:
            if row.get("Componente GEI") is not None:
                comp = row.get("Nombre del Componente")
                pcg = row.get("Potencial de Calentamiento Global")
                if comp and pcg is not None:
                    gei_dict[comp] = pcg
        return gei_dict

    def calcular_emisiones(self, consumos):
        """
        Calcula las emisiones GEI en función de los consumos ingresados por el usuario.
        consumos: dict { nombre_componente: (valor, unidad) }

        Ejemplo:
        consumos = {
           "CO2": (1000, "g"),  # 1000 g de CO2
           "CH4": (2, "lb")     # 2 lb de CH4
        }

        Retorna un dict {nombre_componente: emisiones_kgCO2e}
        """
        resultados = {}
        for sustancia, (valor, unidad) in consumos.items():
            # Verificar si la sustancia es GEI:
            if sustancia not in self.gei_info:
                # No está en GEI, ignorar o lanzar error
                # Podríamos lanzar una advertencia, pero aquí solo ignoramos
                continue
            
            # Convertir el consumo a unidad base (kg)
            valor_kg = self.gestor_unidades.convertir(valor, unidad, self.unidad_base)

            # Obtener PCG
            pcg = self.gei_info[sustancia]
            # Emisiones = consumo_en_kg * PCG
            emisiones = valor_kg * pcg
            

            resultados[sustancia] = emisiones
        return resultados