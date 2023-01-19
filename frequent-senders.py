import mailbox
import subprocess
from tqdm import tqdm
from collections import Counter
from email.parser import BytesParser
from email.policy import compat32
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


# domains = second level domains (facebook.com, ebay.com etc)
domains = []

# subdomains = second level domain + all subdomains (mail.service.facebook.com)
subdomains = []

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


def parseinbox():
    INBOX = "C:/Users/WDAGUtilityAccount/Desktop/INBOX"
    mymail = mailbox.mbox(INBOX, factory=BytesParser(policy=compat32).parse)
    print(
        "\nProcessing inbox file... \n(a progress bar will appear once the mailbox has been loaded into memory)"
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
                subdomain = address.rsplit("@", 1)[1]
                subdomains.append(subdomain)
                domainsplit = subdomain.rsplit(".", 2)
                domain = domainsplit[-2] + "." + domainsplit[-1]
                domains.append(domain)
            except:
                continue


def open_pdf(pdf_path):
    edge = subprocess.Popen(
        [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            pdf_path,
        ], creationflags=subprocess.DETACHED_PROCESS, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT
    )


def bar_chart(data, title, xlabel, ylabel, pdf):
    most_common_list = Counter(data).most_common()[:10]
    # Create the bar chart
    plt.barh(*zip(*most_common_list))
    # Add labels and title
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.tight_layout()
    # Save the chart
    pdf.savefig()
    plt.close()


parseinbox()

write_list_to_csv(domains, "domains.csv")
write_list_to_csv(subdomains, "subdomains.csv")
write_list_to_csv(addresses, "addresses.csv")

with PdfPages("report.pdf") as pdf:
    bar_chart(domains, "Most common domains", "Domain", "Count", pdf)
    bar_chart(subdomains, "Most common subdomains", "Domain", "Count", pdf)
    bar_chart(addresses, "Most common email addresses",
              "Address", "Count", pdf)

open_pdf("report.pdf")
