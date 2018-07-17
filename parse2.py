import csv, json, xlrd

def csvParse(csvfile):
    # Open the CSV
    f = open( csvfile, 'rU' )
        # Change each fieldname to the appropriate field name. 
    reader = csv.DictReader( f, fieldnames = ( "Protocol","Source Port","Destination Port","Source IP", "Destination IP", "Source MAC", "Destination MAC", "segment ID", "Frame name", "Frame ID", "segment name", "packet ID", "Receivers"))
    framenames = []
    store = []
    # Store frame names in a list
    for row in reader:
        frame= {"FrameName":row["Frame name"], 
        "FrameID": row["Frame ID"],
        "protocol": row["Protocol"],
        "segments":[]}
        if row["Frame name"] not in framenames:
            framenames.append(row["Frame name"])
            store.append(frame)
   
    # Create Objects for Frames, segments and packets
    segment = {"segmentName":""}
    for frame in store:
        f = open( csvfile, 'rU' )
        reader = csv.DictReader( f, fieldnames = ( "Protocol","Source Port","Destination Port","Source IP", "Destination IP", "Source MAC", "Destination MAC", "segment ID","Frame name", "Frame ID", "segment name", "packet ID", "Receivers"))
        for row in reader:
            if frame["FrameName"] == row["Frame name"]:
                if segment["segmentName"] != row["segment name"]:
                    segment = {
                        "segmentID":row["segment ID"],
                        "segmentName": row["segment name"],
                        "srcPort":row["Source Port"],
                        "destPort":row["Destination Port"],
                        "packets":[{
                            "packetID":row["packet ID"],
                            "Receivers":row["Receivers"],
                            "destIP":row["Destination IP"],
                            "destMAC":row["Destination MAC"],
                            "srcIP": row["Source IP"],
                            "srcMAC":row["Source MAC"]
                            }]
                        }
                    frame["segments"].append(segment)
                else:
                    packet = {
                        "packetID":row["packet ID"],
                        "Receivers":row["Receivers"],
                        "destIP":row["Destination IP"],
                        "destMAC":row["Destination MAC"],
                        "srcIP": row["Source IP"],
                        "srcMAC":row["Source MAC"]
                        }
                    segment["packets"].append(packet)

    # Parse the CSV into JSON
    out = json.dumps(store, indent=4)
    # Save the JSON
    f = open( 'data.json', 'w')
    f.write(out)


def xlsParse(xlsfile):
    store = []  #List of objects to be parsed into json file
    framenames = []
    book = xlrd.open_workbook(xlsfile)
    sh1 = book.sheet_by_index(0)

    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[8].value not in framenames:
            framenames.append(sh1.row(rx)[8].value)
            frame = {"frameName": sh1.row(rx)[8].value,
               "frameID": sh1.row(rx)[9].value,
               "protocol": sh1.row(rx)[0].value,
               "segments":[]
               }
            store.append(frame)

    segment = {"segmentName":""}
    for frame in store:
        for rx in range(1, sh1.nrows):
            if frame["frameName"] == sh1.row(rx)[8].value:
                if segment["segmentName"] != sh1.row(rx)[10].value:
                    segment = {
                    "segmentID":sh1.row(rx)[7].value,
                    "segmentName": sh1.row(rx)[10].value,
                    "srcPort":int(sh1.row(rx)[1].value),
                    "destPort":int(sh1.row(rx)[2].value),
                    "packets":[{
                        "packetID":sh1.row(rx)[11].value,
                        "Receivers":sh1.row(rx)[12].value,
                        "srcMAC": sh1.row(rx)[5].value,
                        "srcIP": sh1.row(rx)[3].value,
                        "destIP":sh1.row(rx)[4].value,
                        "destMAC":sh1.row(rx)[6].value
                        }]
                    }
                    frame["segments"].append(segment)
                else:
                    packet = {
                    "packetID":sh1.row(rx)[11].value,
                    "Receivers":sh1.row(rx)[12].value,
                    "srcMAC": sh1.row(rx)[5].value,
                    "srcIP": sh1.row(rx)[3].value,
                    "destIP":sh1.row(rx)[4].value,
                    "destMAC":sh1.row(rx)[6].value
                    }
                    segment["packets"].append(packet)

    # Parse the XLS into JSON
    out = json.dumps(store, indent=4)
    # Save the JSON
    f = open( 'xlsdata.json', 'w')
    f.write(out)


csvParse('dataset.csv')

xlsParse('dataset.xls')

