import json
import datetime
import paho.mqtt.subscribe as subscribe
from influxdb import InfluxDBClient
client = InfluxDBClient('localhost', 8086, 'root', 'root', 'Savina')
client.create_database('Savina')
topics = ["SAVINA/SPMS10kW/27HOANGHOATHAM/DANANG/07052020/GateWay01"]
auth = {"username":"mind", "password":"123"}
hostname = "202.158.245.111"
b = {}
while True:
    m = subscribe.simple(topics, transport="tcp", hostname=hostname, port=16766, auth=auth)
    m_payload_dict = json.loads(m.payload)
    ts_noTZ = datetime.datetime.strptime(str(m_payload_dict['ts'])[0:19], "%Y-%m-%dT%H:%M:%S")
    for i in m_payload_dict['d']:
        json_body = [
        {
            "measurement":'GateWay01',
            "tags":
            {
                "tag":i['tag'],
                "year":str(ts_noTZ.year),
                "month":str(ts_noTZ.month),
                "day":str(ts_noTZ.day),
                "hour":str(ts_noTZ.hour), 
		"geohash":"w6ugq76r5",
		"Station":"377 Nguyễn Văn Linh"
            },
            "fields":
            {
                "value":i['value']
            }
        }
        ]
        client.write_points(json_body)