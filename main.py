import subprocess
import msal
import os
from dotenv import load_dotenv


load_dotenv()


class EMSWrapper:

    def __init__(self):
        self.connect()

    def get_token(self):
        app = msal.ConfidentialClientApplication(
            client_id=os.getenv("CLIENT_ID"),
            authority=f"https://login.microsoftonline.com/{os.getenv('TENANT_ID')}",
            client_credential=os.getenv("CLIENT_SECRET"),
        )

        return app.acquire_token_for_client(
            scopes=["https://outlook.office365.com/.default"]
        )["access_token"]

    def read_output(self, ps_session):
        output = ps_session.stdout.read(1)
        while ps_session.poll() is None or output:
            output += ps_session.stdout.read(1)
            if output.endswith("\n"):
                break
        return output

    def connect(self):
        ps_session = subprocess.Popen(
            ["powershell", "-NoExit", "-Command", "-"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        ps_session.stdin.write(
            "Connect-ExchangeOnline -UserPrincipalName chief_kenny_officer@huskerly.onmicrosoft.com -AccessToken {}\n".format(
                self.get_token()
            )
        )
        ps_session.stdin.flush()
        output = ps_session.stdout.read(887)
        print(output)
        # print(self.read_output(ps_session))

        # ps_session.stdin.write("Get-User\n")
        # ps_session.stdin.flush()
        # print(self.read_output(ps_session))

        ps_session.stdin.write("exit\n")
        ps_session.stdin.flush()

        print(ps_session.communicate())


if __name__ == "__main__":
    wrapper = EMSWrapper()
