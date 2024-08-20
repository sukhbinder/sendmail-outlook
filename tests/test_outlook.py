import sys
from unittest import mock
from send_outlookemail import create_parser


# def test_main(tmp_path):
#     with mock.patch("win32com.client.Dispatch") as dispatch_mock:
#         dispatch_mock.return_value.CreateItem.return_value = "MockMail"

#         with mock.patch("os.path.join", return_value=str(tmp_path / "email.msg")):
#             args = [
#                 "sendmail",
#                 "-t",
#                 "test@example.com",
#                 "-s",
#                 "Test Subject",
#                 "-b",
#                 "Test body 1, Test body 2",
#                 "-hb",
#                 "<html><body>HTML Test</body></html>",
#                 "-ds"
#             ]
#             sys.argv = args
    
#             main()

#         assert dispatch_mock.called

def test_create_parser():
    paser = create_parser()
                args = [
                "sendmail",
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
    args = parser.parse_args(args)
    assert args.to == "test@example.com"
    assert args.dont_send
    assert args.subject == "Test Subject"
