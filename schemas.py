from pydantic import BaseModel, Field
from typing import List, Optional
   # Subclase de propiedades, que incluya el valor y las unidades 
class Propiedad(BaseModel):
    valor: float = Field(..., alias='Valor', description='Valor de la propiedad')
    unidades: str = Field(..., alias='Unidades', description='Unidades de la propiedad')
class indicaciones(BaseModel):
    codigo: str = Field(..., alias='Código', description='Código de la indicación de peligro o palabra de advertencia')
    descripcion: str = Field(..., alias='Descripción', description='Indicación de peligro o palabra de advertencia')
class componente(BaseModel):
    nombre: str = Field(..., alias='Nombre', description='Nombre del componente')
    numero_cas: str = Field(..., alias='Número CAS', description='Número CAS del componente')
    porcentaje: str = Field(..., alias='Porcentaje', description='Porcentaje de composición del componente')

class HDSData(BaseModel):
    nombre_sustancia_quimica: str = Field(..., alias='Nombre de la Sustancia Química', description='Nombre de la sustancia química en Español, se encuentra en la Sección 1 de la HDS')
    idioma: Optional[str] = Field(None, alias='Idioma de la HDS', description='Idioma en el que se encuentra la HDS, aunque el idioma no sea español, la información extraída será devuelta en español')
    componentes: List[componente] = Field(..., alias='Componentes', description='Componentes de la sustancia química en Español, se encuentran en la Sección 3 de la HDS')
    olor: Optional[str] = Field(None, alias='Olor', description='Olor de la sustancia química en Español, se encuentra en la Sección 9 de la HDS')
    color: Optional[str] = Field(None, alias='Color', description='Color de la sustancia química en Español, se encuentra en la Sección 9 de la HDS')
    velocidad_evaporacion: Optional[Propiedad] = Field(None, alias='Velocidad de evaporación', description='Velocidad de evaporación de la sustancia química, se encuentra en la Sección 9 de la HDS')
    ph: Optional[float] = Field(None, alias='pH de la sustancia', description='Valor de pH de la sustancia química, se encuentra en la Sección 9 de la HDS')
    temperatura_ebullicion: Optional[Propiedad]= Field(None, alias='Temperatura de ebullición', description='Temperatura de ebullición de la sustancia química, se encuentra en la Sección 9 de la HDS')
    punto_congelacion: Optional[Propiedad] = Field(None, alias='Punto de congelación', description='Punto de congelación de la sustancia química, se encuentra en la Sección 9 de la HDS')
    densidad: Optional[Propiedad] = Field(None, alias='Densidad', description='Densidad de la sustancia química, se encuentra en la Sección 9 de la HDS')
    punto_inflamacion: Optional[Propiedad] = Field(None, alias='Punto de inflamación', description='Punto de inflamación de la sustancia química, se encuentra en la Sección 9 de la HDS')
    limite_inf_inflamabilidad: Optional[float] = Field(None, alias='Límite inferior de inflamabilidad', description='Límite inferior de inflamabilidad de la sustancia química, se encuentra en la Sección 9 de la HDS')
    limite_sup_inflamabilidad: Optional[float] = Field(None, alias='Límite superior de inflamabilidad', description='Límite superior de inflamabilidad de la sustancia química, se encuentra en la Sección 9 de la HDS')
    presion_vapor: Optional[Propiedad] = Field(None, alias='Presión de vapor', description='Presión de vapor de la sustancia química, se encuentra en la Sección 9 de la HDS')
    solubilidad_agua: Optional[Propiedad] = Field(None, alias='Solubilidad en agua', description='Solubilidad en agua de la sustancia química, se encuentra en la Sección 9 de la HDS')
    propiedades_explosivas: Optional[str] = Field(None, alias='Propiedades Explosivas', description='Propiedades explosivasde la sustancia química, se pueden encontrar en la Secciónes 4, 8 y 9  de la HDS')
    propiedades_comburentes: Optional[str] = Field(None, alias='Propiedades Comburentes', description='Propiedades comburentes de la sustancia química, se pueden encontrar en la Secciónes 4, 8 y 9  de la HDS')
    tamano_particula: Optional[str] = Field(None, alias='Tamaño de partícula', description='Tamaño de partícula de la sustancia química, se encuentra en la Sección 9 de la HDS')
    indicaciones_toxicologia: Optional[str] = Field(None, alias='Indicaciones de toxicología', description='Indicaciones de toxicología de la sustancia química, se encuentran en la Sección 11 de la HDS')
    palabra_advertencia: Optional[str] = Field(None, alias='Palabra de Advertencia', description='Palabra de advertencia de la sustancia química conforme a la NOM-018-STPS-2015 SOLO PUEDE SER PELIGRO O ATENCIÓN se encuentra en la Sección 2 de la HDS')
    identificaciones_peligro_h: List[indicaciones] = Field(..., alias='Identificaciones de peligro H', description='Identificaciones de peligro H de la sustancia química, se encuentran en la Sección 2 de la HDS')
    consejos_prudencia_p: List[indicaciones] = Field(..., alias='Consejos de Prudencia P', description='Consejos de prudencia P de la sustancia química, se encuentran en la Sección 2 de la HDS')
    pictogramas: List[str] = Field(..., alias='Pictogramas', description='Pictogramas de la sustancia química, se encuentran en la Sección 2 de la HDS, de igual manera puedes inferirlos con base al Sistema Globalmente armonizado y las indicaciones de peligror H')

