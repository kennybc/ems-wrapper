import subprocess
import os
from threading import Thread
from dotenv import load_dotenv

load_dotenv()


class EMSWrapper:

    def __init__(self):
        self.connect()

    def __del__(self):
        self.disconnect()

    def connect(self):
        self.ps_session = subprocess.Popen(
            ["powershell"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        self.ps_session.stdin.write(
            "Connect-ExchangeOnline -CertificateThumbPrint {} -AppID {} -Organization {}\n".format(
                os.getenv("CERT_THUMB"),
                os.getenv("APP_ID"),
                "Huskerly.onmicrosoft.com",
            )
        )
        self.ps_session.stdin.flush()
        print(self.ps_session.stdout.read(1247))

        print("connected")

    def disconnect(self):
        self.ps_session.stdin.write("Disconnect-ExchangeOnline -Confirm:$false")
        self.ps_session.stdin.flush()

        self.ps_session.stdin.write("exit\n")
        self.ps_session.stdin.flush()

        print(self.ps_session.communicate())

    def read_output(self, output):
        result = ""
        while self.ps_session.poll() is None:
            next = self.ps_session.stdout.read(1)
            print(repr(next))
            if next == "\n":
                print("ends with new line")
                break
            result += next
            if result.endswith("\n"):
                print("ends with new lines")
                break

        output.append(result)

    def invoke(self, command):
        self.ps_session.stdin.write(command)
        self.ps_session.stdin.flush()

        # print(self.ps_session.communicate())

        output = []
        thread = Thread(target=self.read_output, args=(output,))
        # thread.daemon = True
        thread.start()
        print(output)


if __name__ == "__main__":
    wrapper = EMSWrapper()
    wrapper.invoke("echo 'test'")
