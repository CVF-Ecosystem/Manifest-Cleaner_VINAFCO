"""Size/Type mappings for container classification.

This module contains all mapping dictionaries used for converting
container size/type codes to standard formats and cargo type classification.
"""

from typing import Dict


# --- Mapping: Size/Type -> VTOS/ISO ---
# Converts various container size codes to standardized VTOS format.
SIZE_TO_ISO: Dict[str, str] = {
    # Standard Dry Containers
    "20DC": "20DC",
    "20GP": "20DC",
    "20'DC": "20DC",
    "20'GP": "20DC",
    "40DC": "40DC",
    "40GP": "40HC",
    "40'DC": "40DC",
    "40'GP": "40HC",
    "40HC": "40HC",
    "40'HC": "40HC",
    "40HQ": "40HC",
    "40'HQ": "40HC",
    "45GP": "45HC",
    "45HC": "45HC",
    "45'HC": "45HC",
    "45HQ": "45HC",
    
    # ISO Codes Mapping
    "22G1": "20DC",
    "42G1": "40DC",
    "45G1": "40HC",
    "45R1": "45RH", 
    "22R1": "22R0",
    "25G1": "20HC", 
    "42R1": "40RH",
    
    # Reefer Containers
    "20RF": "22R0",
    "20'RF": "22R0",
    "20RH": "22R0",
    "40RF": "40RH",
    "40'RF": "40RH",
    "40RH": "40RH",
    "45RF": "45RH",
    "45RH": "45RH",
    
    # Tank Containers
    "20TK": "20TK",
    "20'TK": "20TK",
    "40TK": "40TK",
    
    # Flat Rack Containers
    "20FR": "20FL",
    "20'FR": "20FL",
    "40FR": "40FL",
    "40'FR": "40FL",
    
    # Open Top Containers
    "20OT": "20OT",
    "20'OT": "20OT",
    "40OT": "40OT",
    "40'OT": "40OT",
}


# --- Mapping: Size/Type -> Base Cargo Type ---
# Used to determine the cargo type based on container size/type.
# GENERAL is the default for standard dry containers.
SIZE_TYPE_TO_CARGO_TYPE_BASE: Dict[str, str] = {
    # Tank
    "20TK": "TANK",
    "20'TK": "TANK",
    "40TK": "TANK",
    
    # General (Dry Containers)
    "20DC": "GENERAL",
    "20GP": "GENERAL",
    "20'DC": "GENERAL",
    "20'GP": "GENERAL",
    "40DC": "GENERAL",
    "40GP": "GENERAL",
    "40'DC": "GENERAL",
    "40'GP": "GENERAL",
    "40HC": "GENERAL",
    "40'HC": "GENERAL",
    "40HQ": "GENERAL",
    "40'HQ": "GENERAL",
    "45HC": "GENERAL",
    "45'HC": "GENERAL",
    "45GP": "GENERAL",
    "45HQ": "GENERAL",
    
    # Reefer
    "20RF": "REEFER",
    "20'RF": "REEFER",
    "20RH": "REEFER",
    "40RF": "REEFER",
    "40'RF": "REEFER",
    "40RH": "REEFER",
    "45RF": "REEFER",
    "45RH": "REEFER",
    
    # Flat Rack
    "20FR": "FLATRACK",
    "20'FR": "FLATRACK",
    "40FR": "FLATRACK",
    "40'FR": "FLATRACK",
    
    # Open Top
    "20OT": "OPENTOP",
    "20'OT": "OPENTOP",
    "40OT": "OPENTOP",
    "40'OT": "OPENTOP",
}


# --- Tare Weights (for F/E determination) ---
# Standard tare weights in TONS for different container sizes.
# Used to determine if a container is Full or Empty based on gross weight.
TARE_WEIGHTS_TONS: Dict[str, float] = {
    "20": 2.3,   # 20ft containers (VGM style)
    "40": 3.8,   # 40ft containers (VGM style)
    "45": 4.5,   # 45ft containers
}

# Reefer container tare weights (heavier due to refrigeration unit)
REEFER_TARE_WEIGHTS_TONS: Dict[str, float] = {
    "20": 3.0,   # 20ft reefer
    "40": 4.8,   # 40ft reefer
}

# Safety margin (in TONS) added to tare weight when determining F/E status
SAFETY_MARGIN_TONS: float = 0.1

# Minimum payload threshold to consider container as FULL
MIN_PAYLOAD_TONS: float = 0.1


# --- Operator Codes ---
OPERATOR_COC = "VFC"  # Carrier Owned Container
OPERATOR_SOC = "SVF"  # Shipper Owned Container
