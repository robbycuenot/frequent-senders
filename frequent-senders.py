import mailbox
from tqdm import tqdm
from collections import Counter
from email.parser import BytesParser
from email.policy import compat32
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


# sldomains = second level domains (facebook.com, ebay.com etc)
sldomains = []

# fulldomains = second level domain + all subdomains (mail.service.facebook.com)
fulldomains = []

# addresses = entire email address (sender12345@mail.service.facebook.com)
addresses = []


def write_list_to_csv(data, filename):
    data_counter = Counter(data)
    most_common_data = data_counter.most_common()
    with open(filename, "w") as csvfile:
        print("Writing to " + filename + "...")
        for row in most_common_data:
            # write each item on a new line
            csvfile.write(row[0] + ", " + str(row[1]))
            csvfile.write("\n")


def readinbox():
    INBOX = "C:/Users/WDAGUtilityAccount/Desktop/INBOX"
    mymail = mailbox.mbox(INBOX, factory=BytesParser(policy=compat32).parse)

    print(
        "\nReading inbox file... \n(a progress bar will appear once the mailbox has been loaded into memory)"
    )

    for _, message in enumerate(tqdm(mymail)):
        messageID = message["Message-ID"]
        if messageID == None:
            continue

        sender = message["from"]
        if sender is not None and isinstance(sender, str):
            try:
                address = sender.rsplit("<")[-1].rstrip().rstrip(">").rstrip()
                addresses.append(address)
                fulldomain = address.rsplit("@", 1)[1]
                fulldomains.append(fulldomain)
                domainsplit = fulldomain.rsplit(".", 2)
                sldomain = domainsplit[-2] + "." + domainsplit[-1]
                sldomains.append(sldomain)
            except:
                continue


readinbox()

write_list_to_csv(sldomains, "sldomains.csv")
write_list_to_csv(fulldomains, "fulldomains.csv")
write_list_to_csv(addresses, "addresses.csv")


data_counter = Counter(sldomains)
most_common_data = data_counter.most_common()
addylist = []
countlist = []

for row in most_common_data[:10]:
    addylist.append(row[0])
    countlist.append(int(row[1]))

with PdfPages("multipage_pdf.pdf") as pdf:
    plt.pie(countlist, labels=addylist, autopct="%.2f%%")
    plt.title("Most frequent senders", fontsize=20)
    pdf.savefig()
    plt.close
