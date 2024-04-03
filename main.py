import pika
from io import BytesIO
import functions.excel_functions.main as planilha


def callback(ch, method, properties, body):
    print('Mensagem em Andamento...')
    with BytesIO(body) as planilha_temporaria:
        result = planilha.pegar_dados_intervalo_planilha(planilha_temporaria, 'A1:E', True)

    if isinstance(properties.reply_to, str):
        """
        Definição das propriedades da resposta: Cria um objeto BasicProperties para a resposta, 
        mantendo a mesma correlation_id da mensagem original.
        """
        response_properties = pika.BasicProperties(correlation_id=properties.correlation_id)

        """
        Envio da resposta para a fila de resposta: Publica a resposta na fila especificada em properties.reply_to, 
        utilizando as propriedades definidas em response_properties e convertendo o resultado para string.
        """
        channel.basic_publish(exchange='', routing_key=properties.reply_to, properties=response_properties,
                              body=str(result))

        """
        Confirmação do processamento da mensagem: Confirma que a mensagem foi processada com sucesso 
        e pode ser removida da fila.
        """
        ch.basic_ack(delivery_tag=method.delivery_tag)
        print('Mensagem Finalizada!')
    else:
        print('Mensagem Finalizada!')
        return


"""
Conexão com o RabbitMQ e configuração do canal: Estabelece a conexão com o RabbitMQ e cria um canal de comunicação.
"""
connection = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
channel = connection.channel()

"""
Declaração da fila de entrada: Declara a fila que será usada para receber mensagens.
"""
channel.queue_declare(queue='other_test')

"""
Consumo de mensagens da fila: Define a função de callback para processar mensagens recebidas da fila 'other_test'.
"""
channel.basic_consume(queue='other_test', on_message_callback=callback)

"""
Início da escuta por mensagens: Inicia o loop de escuta por mensagens na fila. 
O código ficará aguardando novas mensagens e, quando receber uma, chamará a função de callback para processá-la.
"""
print(' [*] Aguardando mensagens. Para sair, pressione CTRL+C')
channel.start_consuming()
