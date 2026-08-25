"""Unit tests for OTP Extractor & Text Sanitizer."""

from __future__ import annotations
import unittest
from datetime import datetime, timezone
from outlook_graph.otp.otp_extractor import (
    clean_html,
    extract_otp,
    extract_all_otps,
    cleanHtml,
    extractOtp,
    extractAllOtps,
)
from outlook_graph.types.mail import ItemBody, OutlookMessage


class TestOtpExtractor(unittest.TestCase):
    """Test suite for multi-language OTP detection & confidence scoring."""

    def test_clean_html_tags_and_decode_entities(self) -> None:
        raw_html = (
            "<div>Your code is <strong>123456</strong>.&nbsp;Do not share.</div>"
        )
        cleaned = clean_html(raw_html)
        self.assertEqual(cleaned, "Your code is 123456 . Do not share.")

    def test_extract_6_digit_code_english_template(self) -> None:
        email = """
        Hi John,
        Your verification code is 739201.
        This code expires in 10 minutes.
        """
        result = extract_otp(email)
        self.assertIsNotNone(result)
        self.assertEqual(result.code, "739201")
        self.assertEqual(result.digits, 6)
        self.assertGreater(result.confidence, 0.85)

    def test_extract_hyphenated_otp_format(self) -> None:
        email = "Your Microsoft verification code is: 492-105"
        result = extract_otp(email)
        self.assertIsNotNone(result)
        self.assertEqual(result.code, "492105")

    def test_extract_4_and_8_digit_otps(self) -> None:
        email4 = "Your login PIN code: 4912"
        res4 = extract_otp(email4)
        self.assertIsNotNone(res4)
        self.assertEqual(res4.code, "4912")

        email8 = "Discord authorization code: 92837401"
        res8 = extract_otp(email8)
        self.assertIsNotNone(res8)
        self.assertEqual(res8.code, "92837401")

    def test_extract_otp_from_vietnamese_email_templates(self) -> None:
        vi1 = "Mã xác thực Shopee của bạn là 810492. Mã này có hiệu lực trong 5 phút. Vui lòng không chia sẻ mã này với bất kỳ ai."
        res1 = extract_otp(vi1)
        self.assertIsNotNone(res1)
        self.assertEqual(res1.code, "810492")

        vi2 = "Mã OTP MoMo: 301928 là mã bảo mật để xác nhận đăng nhập."
        res2 = extract_otp(vi2)
        self.assertIsNotNone(res2)
        self.assertEqual(res2.code, "301928")

        vi3 = "Ma xac nhan: 582910. Khong cung cap ma OTP cho nguoi khac."
        res3 = extract_otp(vi3)
        self.assertIsNotNone(res3)
        self.assertEqual(res3.code, "582910")

    def test_extract_otp_directly_from_outlook_message(self) -> None:
        msg = OutlookMessage(
            id="123",
            subject="Steam Guard Code: 83921",
            body_preview="Here is your Steam verification code",
            body=ItemBody(
                content_type="html",
                content="""
                <html>
                  <body>
                    <p>Hello,</p>
                    <p>Here is your Steam Guard code:</p>
                    <div style="font-size: 24px; font-weight: bold;">83921</div>
                  </body>
                </html>
                """,
            ),
            received_date_time=datetime.now(timezone.utc).isoformat(),
            has_attachments=False,
            is_read=False,
        )
        result = extract_otp(msg)
        self.assertIsNotNone(result)
        self.assertEqual(result.code, "83921")

    def test_ignore_copyright_years_when_scoring(self) -> None:
        email_with_footer = """
        Your security code is 638291.
        Copyright (c) 2026 Company Inc. All rights reserved.
        """
        result = extract_otp(email_with_footer)
        self.assertIsNotNone(result)
        self.assertEqual(result.code, "638291")

    def test_typescript_parity_aliases(self) -> None:
        res = extractOtp("Code: 998877")
        self.assertIsNotNone(res)
        self.assertEqual(res.code, "998877")
        self.assertEqual(cleanHtml("<b>hello</b>"), "hello")


if __name__ == "__main__":
    unittest.main()
