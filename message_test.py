# CÓDIGO PARA TESTAR O CÓDIGO MAIN:

import pika

# Conexão com o RabbitMQ
connection = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
channel = connection.channel()

# Declaração da fila
channel.queue_declare(queue='test_queue')

# Publica mensagens na fila
for i in range(5):
    message = f'Mensagem de teste {i}'
    channel.basic_publish(exchange='', routing_key='test_queue', body=message)
    print(f' [x] Enviado "{message}"')

# Fecha a conexão
connection.close()
