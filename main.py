import win32com.client
import os
import json
from datetime import date



#creacion de clase
class Message:
    ticket='SAC-202009170001'
    tipooperacion='SAC'
    date='20201010'
    agente=''


#crear objeto .
msg = Message()
msg.ticket='SAC-202009170001'
msg.tipooperacion='CAJA'
msg.date='20201010'
msg.agente=''


#Conexion a cola
qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name= os.getenv('COMPUTERNAME')
qinfo.FormatName="direct=os:" + computer_name + "\\PRIVATE$\\bankMessage"
queue=qinfo.Open(2,0)

#Mensaje
msgq=win32com.client.Dispatch("MSMQ.MSMQMessage")
msgq.Label = "Mensaje de Banco"
msgq.Body= json.dumps(msg.__dict__)
msgq.Send(queue)
print('Mensaje encolado exitosamente a ' + qinfo.FormatName)
queue.Close()

queue=qinfo.Open(1,0)
msgReadJson=queue.Receive()
print('Desencolando mensaje')
print(str(msgReadJson))

#Convertir de json a objeto.
class Deserialize(object):
    def __init__(self, j):
        self.__dict__ = json.loads(j)

p = Deserialize(str(msgReadJson))
print(p.ticket)




