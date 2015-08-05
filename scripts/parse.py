import unicodecsv
import xlrd

SOURCE = "source/EITI Data compilation.xlsx"
LOCATIONS = "source/locations.csv"
OUTPUT = "../data/data.csv"
START_ROW = 6
END_ROW = 36 # Hard-coded for now
COL_MAPPINGS = {
    2: "project_name",
    3: "tax_id",
    4: "commodity",
    5: "location",
    6: "region",
    7: "concession_size",
    10: "production_2006",
    11: "production_2007",
    12: "production_2008",
    13: "production_2009",
    14: "production_2010",
    15: "production_2011",
    16: "production_2012",
    17: "production_2013",
    # Extend this when revenue data available
}
OUTPUT_COLS = ['project_name', 'tax_id', 'commodity', 'location',
     'region', 'concession_size', 'production_vol', 'year', 'lat', 'long']

def run():
    of = open(OUTPUT, "wb")
    oc = unicodecsv.DictWriter(of, fieldnames=OUTPUT_COLS)
    oc.writeheader()

    # Read in locations and store lat and long aside them in a dict
    lf = open(LOCATIONS, "r")
    lc = unicodecsv.DictReader(lf)
    locs = dict(map(lambda x: (x['location'], ({
        'lat': x['lat'], 'long': x['long']})
        ), lc))

    bk = xlrd.open_workbook(SOURCE)
    st = bk.sheet_by_name("Project Level Production")
    
    # Get row data for each relevant row
    for rownum in range(START_ROW, END_ROW):
        row_data = dict(map(lambda x: (x[1], st.cell_value(rownum, x[0])),
                     COL_MAPPINGS.items()))
        if row_data["project_name"] in ["", "Oil and Gas Projects"]:
            continue

        # Add location data to row data
        row_loc = locs.get(row_data['location'], "")
        row_data.update(row_loc)
        d = dict([(col, row_data.get(col)) for col in OUTPUT_COLS])

        def filter_prodn(val):
            return val.startswith("production_")

        # Make one row per year, showing production vol per year
        prodn = filter(filter_prodn, COL_MAPPINGS.values())
        pys = list([(p, p.split("_")) for p in prodn])
        for lookup, vals in pys:
            d['year'] = vals[1]
            d['production_vol'] = row_data.get(lookup)
            oc.writerow(d)
    of.close()
    lf.close()

if __name__ == '__main__':
    run()
