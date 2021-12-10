import sys
from pathlib import Path

import numpy as np
import dlite
from dlite.utils import instance_from_dict

sys.path.append(str(Path.home() / 'prosjekter/EMMC/OntoTrans/Interfaces/oteapi'))
#import app


# Set up some paths
thisdir = Path(__file__).parent

#config = app.models.resourceconfig.ResourceConfig(
#    downloadUrl=('https://file-examples-com.github.io/uploads/2017/02/'
#                 'file_example_XLSX_10.xlsx'),
#    mediaType=('application/vnd.openxmlformats-officedocument.'
#               'spreadsheetml.sheet'),
#    configuration = {
#        "worksheet": 'Sheet1',
#        "row_to": 10,
#        "col_to": 8,
#        "header_positions": ["A1", "B1", "C1", "D1", "E1", "F6", "G1", "H1"],
#    },
#)
#parser = app.strategy.iparsestrategy.create_parse_strategy(config)
#output = parser.parse()


# Quick load into numpy structured array
import pandas as pd
rawdata = pd.read_excel(thisdir / 'rawdata.xlsx').to_numpy()
data = np.rec.fromrecords(
    rawdata[1:],
    names='Sample,RunID,EIS,limp24h,imp24h,limp2h,imp2h,'
    'Ecorr,icorr,Beta_a,Beta_c,error,'
    'LPR30min,LPR1h,LPR2h,LPR3h,LPR6h,LPR12h,LPR18h,LPR24h')

datamodel = {
    "uri": "http://vipcoat.eu/meta/0.1/sample",
    "meta": "http://onto-ns.com/meta/0.3/EntitySchema",
    "description": "Datamodel for a sample",
    "dimensions": [
        {"name": "nruns", "description": "Number of runs."},
        {"name": "nimp", "description": "Number of impedance measurements."},
        {"name": "nlpr", "description": "Number of LPR measurements"}
    ],
    "properties": [
        {
            "name": "Sample",
            "type": "string",
            "description": "Sample name",
        },
        {
            "name": "RunID",
            "type": "string",
            "dims": ["nruns"],
            "description": "ID for the different runs.",
        },
        {
            "name": "date",
            "type": "string",
            "dims": ["nruns"],
            "description": "Date of each run.",
        },
        {
            "name": "composition",
            "type": "string",
            "dims": ["nruns"],
            "description": "Composition of each run.",
        },
        {
            "name": "substrate",
            "type": "string",
            "dims": ["nruns"],
            "description": "The substrate for each run.",
        },
        {
            "name": "substrate",
            "type": "string",
            "dims": ["nruns"],
            "description": "The substrate for each run.",
        },
        {
            "name": "impedance_time",
            "type": "float64",
            "dims": ["nimp"],
            "unit": "h",
            "description": "Duration of the different impedance measurements.",
        },
        {
            "name": "log_impedance",
            "type": "float64",
            "dims": ["nruns", "nimp"],
            "unit": "log(Ohm)",
            "description": "Logarithm of measured impedance for each run and time.",
        },
        {
            "name": "impedance",
            "type": "float64",
            "dims": ["nruns", "nimp"],
            "unit": "Ohm",
            "description": "Measured impedance for each run and time.",
        },
        {
            "name": "Ecorr",
            "type": "float64",
            "dims": ["nruns"],
            "unit": "mV",
            "description": "Measured corrosion voltage for each run.",
        },
        {
            "name": "icorr",
            "type": "float64",
            "dims": ["nruns"],
            "unit": "uA",
            "description": "Measured corrosion current for each run.",
        },
        {
            "name": "Beta_a",
            "type": "float64",
            "dims": ["nruns"],
            "unit": "mV",
            "description": "Beta_a for each run.",
        },
        {
            "name": "Beta_c",
            "type": "float64",
            "dims": ["nruns"],
            "unit": "mV",
            "description": "Beta_c for each run.",
        },
        {
            "name": "fit_error",
            "type": "float64",
            "dims": ["nruns"],
            "description": "Fit error for each run.",
        },
        {
            "name": "LPR_time",
            "type": "float64",
            "dims": ["nlpr"],
            "unit": "h",
            "description": "Time for each LPR measurement.",
        },
        {
            "name": "LPR",
            "type": "float64",
            "dims": ["nruns", "nlpr"],
            "description": "LPR for each run and time.",
        }
    ]
}
meta = instance_from_dict(datamodel)
