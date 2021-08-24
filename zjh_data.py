from bs4 import BeautifulSoup
import requests
import re
from urllib.parse import urljoin

zjh_url = "http://www.csrc.gov.cn/pub/zjhpublic/G00306205/201711/t20171106_326522.htm"


def download_zjh_xls(path):
    resp = requests.get(zjh_url)
    resp.encoding = "utf-8"
    soup = BeautifulSoup(resp.text)
    print(resp.text)
    node = soup.find_all(href=re.compile(r"xls"))
    print(node[0]['href'])
    xls_url = urljoin(zjh_url, node[0]['href'])
    with open(path, "wb") as f:
        resp = requests.get(xls_url)
        f.write(resp.content)


if __name__ == "__main__":
    download_zjh_xls("zjh_test.xls")
