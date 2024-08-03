import sys
from unittest import mock
from send_outlookemail import main


def test_main(tmp_path):
    with mock.patch("win32com.client.Dispatch") as dispatch_mock:
        dispatch_mock.return_value.CreateItem.return_value = "MockMail"

        with mock.patch("os.path.join", return_value=str(tmp_path / "email.msg")):
            args = [
                "-t",
                "test@example.com",
                "-s",
                "Test Subject",
                "-b",
                "Test body 1, Test body 2",
                "-hb",
                "<html><body>HTML Test</body></html>",
                "-ds"
            ]
            sys.argv = args
    
            main()

        assert dispatch_mock.called