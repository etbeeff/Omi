import mailbox

import os

import sys

import traceback



from exchangelib import Account, Configuration, Credentials, DELEGATE





USERNAME = 'abhinav@ericsson.com'

PASSWORD = 'Vineeta@09031979'

SERVER = 'outlook.office365.com'



ID_FILE = '.read_ids'





def create_mailbox_message(e_msg):

    m = mailbox.mboxMessage(e_msg.mime_content)

    if e_msg.is_read:
        print("A")
        m.set_flags('S')

    return m



def get_read_ids():

    if os.path.exists(ID_FILE):

        with open(ID_FILE, 'r') as f:
            print("B")

            return set([s for s in f.read().splitlines() if s])

    else:
        print("C")
        return set()



def set_read_ids(ids):

    with open(ID_FILE, 'w') as f:

        for i in ids:

            if i:
                print("D")
                f.write(i)

                f.write(os.linesep)

    



if __name__ == '__main__':
    print("Abhi",sys.argv)
    if len(sys.argv) != 3:
        print("E")

        print("Usage: {} folder_name mbox_file".format(sys.argv[0]))

        sys.exit()

    credentials = Credentials(USERNAME, PASSWORD)

    config = Configuration(server=SERVER, credentials=credentials)

    account = Account(primary_smtp_address=USERNAME, config=config, autodiscover=False, access_type=DELEGATE)

    mbox = mailbox.mbox(sys.argv[2])
    print("omi",mbox)

    mbox.lock()

    read_ids_local = get_read_ids()

    folder = getattr(account, sys.argv[1], None)

    item_ids_remote = list(folder.all().order_by('-datetime_received').values_list('item_id', 'changekey'))

    total_items_remote = len(item_ids_remote)

    new_ids = [x for x in item_ids_remote if x[0] not in read_ids_local]

    read_ids = set()

    print("Total items in folder {}: {}".format(sys.argv[1], total_items_remote))

    for i, item in enumerate(account.fetch(new_ids), 1):

        try:

            msg = create_mailbox_message(item)

            mbox.add(msg)

            mbox.flush()

        except Exception as e:

            traceback.print_exc()

            print("[ERROR] {} {}".format(item.datetime_received, item.subject))

        else:

            if item.item_id:

                read_ids.add(item.item_id)

            print("[{}/{}] {} {}".format(i, len(new_ids), str(item.datetime_received), item.subject))

    mbox.unlock()

    set_read_ids(read_ids_local | read_ids)
