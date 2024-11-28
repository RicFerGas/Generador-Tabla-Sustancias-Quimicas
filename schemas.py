from pydantic import BaseModel, Field
from typing import List, Optional
   # Subclase de propiedades, que incluya el valor y las unidades 
class Propiedad(BaseModel):
    valor: float = Field(...,
                         alias='Valor',
                         description='Valor de la propiedad')
    unidades: str = Field(...,
                          alias='Unidades',
                          description='Unidades de la propiedad')
    
class indicaciones(BaseModel):
    codigo: str = Field(..., alias='Código',
                        description='Código de la indicación de peligro o palabra de advertencia')
    descripcion: str = Field(..., 
                             alias='Descripción', 
                             description='Indicación de peligro o palabra de advertencia')
class EstadoFisico(BaseModel):
    solido_baja: bool = Field(..., alias='Sólido de media volatilidad',
                               description="""Sustancias sólidas cristalinas o granulares. 
                               Cuando son usadas, se observa producción de polvo que se disipa o deposita rápidamente sobre superficies después del uso. 
                               p.ej. jabón en polvo, entre otros.""")
    solido_baja: bool = Field(..., alias='Sólido de baja volatilidad',
                                 description="""Sustancias en forma de pellets que no tienen tendencia a romperse.
                                   No se aprecia producción de polvo durante su empleo. 
                                   p.ej. pellets de cloruro de polivinilo, escamas enceradas, entre otras.""")
    solido_alta: bool = Field(..., alias='Sólido de alta volatilidad',
                                description="""Polvos finos, ligeros y de baja densidad. Cuando son usados, se
                                producen nubes de polvo que permanecen en el aire durante varios
                                minutos. p.ej. cemento, negro de humo, polvo de tiza, entre otros.""")
    liquido: bool = Field(..., alias='Líquido', description='Indica si la sustancia química es líquida')
    gaseoso: bool = Field(..., alias='Gaseoso', description='Indica si la sustancia química es gaseosa')

class componente(BaseModel):
    nombre: str = Field(..., 
                        alias='Nombre',
                        description='Nombre del componente')
    numero_cas: str = Field(...,
                            alias='Número CAS',
                            description='Número CAS del componente')
    porcentaje: str = Field(..., alias='Porcentaje',
                            description='Porcentaje de composición del componente')
class ValoresLimiteExposicion(BaseModel):
    oral: Optional[Propiedad] = Field(..., alias='Oral', description='Valor límite de exposición oral')
    inhalacion: Optional[Propiedad] = Field(..., alias='Inhalación', description='Valor límite de exposición por inhalación')
    cutanea: Optional[Propiedad] = Field(..., alias='Cutánea', description='Valor límite de exposición cutánea')

class Pictogramas(BaseModel):
    bomba_explotando: bool = Field(
        ...,
        alias='Bomba explotando',
        description=(
            "Indica si el pictograma de bomba explotando está presente, "
            "código del pictograma GHS01 asociado a sustancias con peligros Explosivos."
        )
    )
    llama: bool = Field(
        ...,
        alias='Llama',
        description=(
            "Indica si el pictograma de llama está presente, "
            "código del pictograma GHS02 asociado a sustancias Inflamables."
        )
    )
    llama_sobre_circulo: bool = Field(
        ...,
        alias='Llama sobre un círculo',
        description=(
            "Indica si el pictograma de llama sobre un círculo está presente, "
            "código del pictograma GHS03 asociado a sustancias Comburentes."
        )
    )
    cilindro_de_gas: bool = Field(
        ...,
        alias='Cilindro de gas',
        description=(
            "Indica si el pictograma de cilindro de gas está presente, "
            "código del pictograma GHS04 asociado a Gases comprimidos."
        )
    )
    corrosion: bool = Field(
        ...,
        alias='Corrosión',
        description=(
            "Indica si el pictograma de corrosión está presente, "
            "código del pictograma GHS05 asociado a sustancias Corrosivas."
        )
    )
    calavera_tibias_cruzadas: bool = Field(
        ...,
        alias='Calavera y tibias cruzadas',
        description=(
            "Indica si el pictograma de calavera y tibias cruzadas está presente, "
            "código del pictograma GHS06 asociado a sustancias con Toxicidad aguda."
        )
    )
    signo_de_exclamacion: bool = Field(
        ...,
        alias='Signo de exclamación',
        description=(
            "Indica si el pictograma de signo de exclamación está presente, "
            "código del pictograma GHS07 asociado a sustancias Irritantes."
        )
    )
    peligro_para_la_salud: bool = Field(
        ...,
        alias='Peligro para la salud',
        description=(
            "Indica si el pictograma de peligro para la salud está presente, "
            "código del pictograma GHS08 asociado a sustancias con Riesgos crónicos para la salud."
        )
    )
    medio_ambiente: bool = Field(
        ...,
        alias='Medio ambiente',
        description=(
            "Indica si el pictograma de medio ambiente está presente, "
            "código del pictograma GHS09 asociado a sustancias con Peligro ambiental."
        )
    )

class HDSData(BaseModel):
    nombre_sustancia_quimica: str = Field(...,
                                          alias='Nombre de la Sustancia Química',
                                          description='Nombre de la sustancia química en Español, se encuentra en la Sección 1 de la HDS')
    idioma: Optional[str] = Field(None,
            alias='Idioma de la HDS',
            description='Idioma en el que se encuentra la HDS, aunque el idioma no sea español, la información extraída será devuelta en español')
    componentes: List[componente] = Field(...,
                                          alias='Componentes',
                                          description='Componentes de la sustancia química en Español, se encuentran en la Sección 3 de la HDS')
    sujeta_retc: Optional[str] = Field(None,
                                       alias='Sujeta a RETC',
                                       description='Elige una de las 3 opciones [NO, SI, REVISAR] Indica si la sustancia química está sujeta o podria estar sujeta al Reglamento de Evaluación de la Conformidad con la NOM-165-SEMARNAT-2013 de México')
    sujeta_gei: Optional[str] = Field(None,
                                      alias='Sujeta a GEI',
                                      description='Elige una de las 3 opciones [NO, SI, REVISAR] Indica si la sustancia química está sujeta o podria estar sujeta a reporte de  Gases de Efecto Invernadero de conformidad con ACUERDO:Qué establece los gases o compuestos de efecto invernadero que se agrupan para efectos de reporte de emisiones, así como sus potenciales de calentamiento en Mexico por la semarnat')
    valoreslimite: Optional[ValoresLimiteExposicion] = Field(..., alias='Valores límite de exposición', description='Valores límite de exposición de la sustancia química, se encuentran en la Sección 8 y 11 de la HDS')
    estado_fisico: Optional[EstadoFisico] = Field(..., alias='Estado Físico', description='Estado físico de la sustancia química en Español, se encuentra en la Sección 9 de la HDS')
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
    peso_molecular: Optional[Propiedad] = Field(None, alias='Peso molecular', description='Peso molecular de la sustancia química, se encuentra en la Sección 9 de la HDS')
    presion_vapor: Optional[Propiedad] = Field(None, alias='Presión de vapor', description='Presión de vapor de la sustancia química, se encuentra en la Sección 9 de la HDS')
    solubilidad_agua: Optional[Propiedad] = Field(None, alias='Solubilidad en agua', description='Solubilidad en agua de la sustancia química, se encuentra en la Sección 9 de la HDS')
    propiedades_explosivas: Optional[str] = Field(None, alias='Propiedades Explosivas', description='Propiedades explosivasde la sustancia química, se pueden encontrar en la Secciónes 4, 8 y 9  de la HDS')
    propiedades_comburentes: Optional[str] = Field(None, alias='Propiedades Comburentes', description='Propiedades comburentes de la sustancia química, se pueden encontrar en la Secciónes 4, 8 y 9  de la HDS')
    tamano_particula: Optional[str] = Field(None, alias='Tamaño de partícula', description='Tamaño de partícula de la sustancia química, se encuentra en la Sección 9 de la HDS')
    indicaciones_toxicologia: Optional[str] = Field(None, alias='Indicaciones de toxicología', description='Indicaciones de toxicología de la sustancia química, se encuentran en la Sección 11 de la HDS')
    palabra_advertencia: Optional[str] = Field(None, alias='Palabra de Advertencia', description='Palabra de advertencia de la sustancia química conforme a la NOM-018-STPS-2015 SOLO PUEDE SER PELIGRO O ATENCIÓN se encuentra en la Sección 2 de la HDS')
    identificaciones_peligro_h: List[indicaciones] = Field(..., alias='Identificaciones de peligro H', description='Identificaciones de peligro H de la sustancia química, se encuentran en la Sección 2 de la HDS')
    consejos_prudencia_p: List[indicaciones] = Field(..., alias='Consejos de Prudencia P', description='Consejos de prudencia P de la sustancia química, se encuentran en la Sección 2 de la HDS')
    pictogramas: Pictogramas = Field(..., alias='Pictogramas', description='Pictogramas de la sustancia química, se encuentran en la Sección 2 de la HDS, de igual manera puedes inferirlos con base al Sistema Globalmente armonizado y las indicaciones de peligror H')

    