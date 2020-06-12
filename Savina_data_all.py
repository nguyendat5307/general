import json
import paho.mqtt.subscribe as subscribe
from influxdb import InfluxDBClient
import datetime

client = InfluxDBClient('localhost', 8086, 'root', 'root', 'SavinaAll')
client.create_database('SavinaAll')
topics = ["SAVINA/SPMS10kW/27HOANGHOATHAM/DANANG/07052020/#"]
auth = {"username":"mind", "password":"123"}
hostname = "202.158.245.111"
b = {}
while True:
    m = subscribe.simple(topics, transport="tcp", hostname=hostname, port=16766, auth=auth, msg_count=5)
    for i in m:
        topic = i.topic
        print(topic[-9:] + ' ' +  str(datetime.datetime.now()))
        payload = json.loads(i.payload)
        for d in payload['d']:
            json_body = [
            {
                "measurement":topic[-9:],
                "tags":
                {
                "tag":d['tag']
                },
                "fields":
                {
                "value":d['value']
                }
            }
            ]
            client.write_points(json_body)
