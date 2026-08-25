import { describe, it, expect } from 'vitest';
import {
  extractOtp,
  extractAllOtps,
  cleanHtml,
} from '../../ts/otp/otp-extractor.js';
import type { OutlookMessage } from '../../ts/types/mail.js';

describe('otp-extractor', () => {
  it('should clean HTML tags and decode entities properly', () => {
    const rawHtml =
      '<div>Your code is <strong>123456</strong>.&nbsp;Do not share.</div>';
    const cleaned = cleanHtml(rawHtml);
    expect(cleaned).toBe('Your code is 123456 . Do not share.');
  });

  it('should extract 6-digit verification code from English template', () => {
    const email = `
Subject: Your GitHub verification code
Body: Hi, please use the following security code to access your account: 739201.
This code expires in 10 minutes.
`;
    const result = extractOtp(email);
    expect(result).not.toBeNull();
    expect(result?.code).toBe('739201');
    expect(result?.digits).toBe(6);
    expect(result?.confidence).toBeGreaterThan(0.85);
  });

  it('should extract hyphenated OTP format (e.g. 123-456)', () => {
    const email = 'Your Microsoft verification code is: 492-105';
    const result = extractOtp(email);
    expect(result).not.toBeNull();
    expect(result?.code).toBe('492105');
  });

  it('should extract 4-digit and 8-digit OTPs', () => {
    const email4 = 'Your login PIN code: 4912';
    expect(extractOtp(email4)?.code).toBe('4912');

    const email8 = 'Discord authorization code: 92837401';
    expect(extractOtp(email8)?.code).toBe('92837401');
  });

  it('should extract OTP from Vietnamese email templates', () => {
    const viEmail1 = `
Mã xác thực Shopee của bạn là 810492. Mã này có hiệu lực trong 5 phút. Vui lòng không chia sẻ mã này với bất kỳ ai.
`;
    const res1 = extractOtp(viEmail1);
    expect(res1?.code).toBe('810492');

    const viEmail2 = `
Mã OTP MoMo: 301928 là mã bảo mật để xác nhận đăng nhập.
`;
    const res2 = extractOtp(viEmail2);
    expect(res2?.code).toBe('301928');

    const viEmail3 = `
Ma xac nhan: 582910. Khong cung cap ma OTP cho nguoi khac.
`;
    const res3 = extractOtp(viEmail3);
    expect(res3?.code).toBe('582910');
  });

  it('should extract OTP directly from an OutlookMessage object', () => {
    const msg: OutlookMessage = {
      id: '123',
      subject: 'Steam Guard Code: 83921',
      bodyPreview: 'Here is your Steam verification code',
      body: {
        contentType: 'html',
        content: `
          <html>
            <body>
              <p>Hello,</p>
              <p>Here is your Steam Guard code:</p>
              <div style="font-size: 24px; font-weight: bold;">83921</div>
            </body>
          </html>
        `,
      },
      receivedDateTime: new Date().toISOString(),
      hasAttachments: false,
      isRead: false,
    };

    const result = extractOtp(msg);
    expect(result?.code).toBe('83921');
  });

  it('should ignore copyright years when scoring codes', () => {
    const emailWithFooter = `
Your security code is 638291.
Copyright (c) 2026 Company Inc. All rights reserved.
`;
    const result = extractOtp(emailWithFooter);
    expect(result?.code).toBe('638291');
  });
});
