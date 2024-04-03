import pika
import uuid

# Conexão com o RabbitMQ
connection = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
channel = connection.channel()

# Declaração da fila de resposta
result = channel.queue_declare(queue='', exclusive=True)
callback_queue = result.method.queue


def on_response(ch, method, properties, body):
    if properties.correlation_id == correlation_id:
        print(f' [x] Resposta recebida: {body}')


# Consumir fila de resposta
channel.basic_consume(queue=callback_queue, on_message_callback=on_response, auto_ack=True)

# Ler o arquivo Excel
with open('teste.xlsx', 'rb') as file:
    file_content = file.read()

# Publica mensagens na fila
correlation_id = str(uuid.uuid4())  # Criando ID unico para a mensagem
channel.basic_publish(
    exchange='',
    routing_key='other_test',
    properties=pika.BasicProperties(
        reply_to=callback_queue,
        correlation_id=correlation_id,
    ),
    body=file_content
)
print(f' [x] Enviado Excel, esperando resposta...')

# Aguardar respostas
connection.process_data_events(time_limit=1)

# Fecha a conexão
connection.close()
