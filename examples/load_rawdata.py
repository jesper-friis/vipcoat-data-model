import sys
import re
from pathlib import Path

import numpy as np
import dlite
from dlite.utils import instance_from_dict

sys.path.append(str(Path.home() / 'prosjekter/EMMC/OntoTrans/Interfaces/oteapi'))
import app


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
            "dims": ["nimp", "nruns"],
            "unit": "log(Ohm)",
            "description": "Logarithm of measured impedance for each run and time.",
        },
        {
            "name": "impedance",
            "type": "float64",
            "dims": ["nimp", "nruns"],
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
            "dims": ["nlpr", "nruns"],
            "description": "LPR for each run and time.",
        }
    ]
}


class UnexpectedExcelFormatError(Exception):
    """Excel file has an unexpected format."""


def get_metadata():
    """Returns excel metadata."""
    return instance_from_dict(datamodel)


def parse_runid(runid):
    """Parse `runid` and return a ``(date, composition, substrate)`` tuple."""
    tokens = runid.split('_')
    if len(tokens) == 3:
        date, composition, _ = tokens
        substrate = ''
    elif len(tokens) == 4:
        date, composition, substrate, _ = tokens
    elif len(tokens) == 5:
        date, comp1, comp2, substrate, _ = tokens
        composition = comp1 + '_' + comp2
    else:
        raise UnexpectedExcelFormatError(f'Unexpected runid: {runid}')
    return date, composition, substrate


def to_seconds(unit):
    """Returns number of seconds that time unit `unit` corresponds to."""
    if unit in ('s', 'second'):
        return 1
    elif unit in ('M', 'min'):
        return 60
    elif unit in ('h', 'hour'):
        return 3600
    elif unit == 'day':
        return 3600 * 24
    elif unit == 'month':
        return 3600 * 24 * 30
    elif unit == 'year':
        return 3600 * 24 * 365.25
    else:
        raise UnexpectedExcelFormatError(f'Unknown time unit: {unit}')


def parse_time_label(label, prefix):
    """Parse (impedance or lpr) label and return the number of hours it encode.
    The label should be of the form

        <prefix><N><unit>

    where
    - <prefix> shold match `prefix`
    - <N> is a number
    - <unit> is a time unit
    """
    m = re.match(
        f'^{prefix}(\d+)(s|second|min|h|hour|day|month|year)s?$', label)
    if not m:
        raise UnexpectedExcelFormatError(
            f'Unexpected impedance label: "{label}". '
            f'Should start with "{prefix}"')
    value, unit = m.groups()
    return float(value) * to_seconds(unit) / 3600


def parse_rawdata_recarray(data : np.recarray) -> dlite.Instance:
    """Parses a rawdata excel sheet represented as a numpy recarray.

    The results are provided as a collection of dlite sample entities.
    """
    coll = dlite.Collection()

    start, = np.nonzero(data.Sample != 'nan')
    end, = np.nonzero(data.RunID == 'AVG')
    if len(start) != len(end):
        raise UnexpectedExcelFormatError(
            f'Number of sample labels (={len(start)}) does not match number '
            f'of RunID "AVG" strings (={len(end)}).')

    imp_labels = [name for name in data.dtype.names if name.startswith('imp')]
    lpr_labels = [name for name in data.dtype.names if name.startswith('LPR')]
    nimp = len(imp_labels)
    nlpr = len(lpr_labels)
    for i, (m, n) in enumerate(zip(start, end)):
        d = data[m:n]
        nruns = n - m
        date, composition, substrate = zip(*[parse_runid(r) for r in d.RunID])
        log_imp = np.zeros((nimp, nruns))
        imp = np.zeros((nimp, nruns))
        for j, label in enumerate(imp_labels):
            log_imp[j] = d['l' + label]
            imp[j] = d[label]
        LPR = np.zeros((nlpr, nruns))
        for j, label in enumerate(lpr_labels):
            LPR[j] = d[label]

        meta = get_metadata()
        inst = meta([nruns, nimp, nlpr])
        inst.Sample = d.Sample[0]
        inst.RunID = d.RunID
        inst.date = date
        inst.composition = composition
        inst.substrate = substrate
        inst.impedance_time = [parse_time_label(label, 'imp')
                               for label in imp_labels]
        inst.log_impedance = log_imp
        inst.impedance = imp
        inst.Ecorr = d.Ecorr
        inst.icorr = d.icorr
        inst.Beta_a = d.Beta_a
        inst.Beta_c = d.Beta_c
        inst.fit_error = d.fit_error
        inst.LPR_time = [parse_time_label(label, 'LPR')
                         for label in lpr_labels]
        inst.LPR = LPR
        coll.add(f'sample{i}', inst)

    return coll


def load_rawdata():
    newHeader = (
        "Sample,RunID,limp24h,imp24h,limp2h,imp2h,"
        "Ecorr,icorr,Beta_a,Beta_c,fit_error,"
        "LPR30min,LPR1h,LPR2h,LPR3h,LPR6h,LPR12h,LPR18h,LPR24h")
    config = app.models.resourceconfig.ResourceConfig(
        #downloadUrl=('https://file-examples-com.github.io/uploads/2017/02/'
        #             'file_example_XLSX_10.xlsx'),
        downloadUrl=('file://rawdata.xlsx'),
        mediaType=('application/vnd.openxmlformats-officedocument.'
                   'spreadsheetml.sheet'),
        configuration = {
            "worksheet": "data",
            "row_from": 3,
            "newHeader": newHeader.split(","),
        },
    )
    parser = app.strategy.iparsestrategy.create_parse_strategy(config)
    output = parser.parse_recarray()
    rec = output["data"]
    return rec


if __name__ == "__main__":
    rec = load_rawdata()
