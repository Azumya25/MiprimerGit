from exchangelib import DELEGATE, Account, Credentials, Configuration

credentials = Credentials(
    #username = 'mario.alberto@cagroup.mx', #or myusername
    username = 'aldo.gallegos@cagroup.mx',
    #password = 'Juc93843'
    password = 'Zub42485'
)

config = Configuration(server='outlook.office365.com', credentials=credentials)

test_account = Account(
    primary_smtp_address = 'aldo.gallegos@cagroup.mx',
    config = config,
    autodiscover = False,
    access_type = DELEGATE
)
# Print first 100 inbox messages in reverse order
for item in test_account.inbox.all().order_by('-datetime_received')[:100]:
    print(item.subject)
    print("hola, cambio 1 para git")
    #print(item.subject, item.body, item.attachments)

def fun_hola():
    print("Otro cambio mas para la rama prueba1")