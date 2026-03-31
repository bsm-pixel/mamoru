/**
 * Gmail 발송 모듈 (Nodemailer)
 * GAS GmailApp.sendEmail() 대체
 *
 * 환경변수:
 * - GMAIL_USER: 발송 Gmail 주소 (bsm@mamoru.kr 또는 Gmail)
 * - GMAIL_APP_PASSWORD: Gmail 앱 비밀번호 (2단계 인증 필요)
 */

import nodemailer from 'nodemailer';

const GMAIL_USER = process.env.GMAIL_USER || '';
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD || '';
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || 'bsm@mamoru.kr';

let transporter: nodemailer.Transporter | null = null;

function getTransporter() {
  if (!transporter && GMAIL_USER && GMAIL_APP_PASSWORD) {
    transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: GMAIL_USER,
        pass: GMAIL_APP_PASSWORD,
      },
    });
  }
  return transporter;
}

export async function sendAdminEmail(subject: string, body: string): Promise<boolean> {
  const t = getTransporter();
  if (!t) {
    console.warn('[email] Gmail 환경변수 미설정 — 이메일 발송 스킵');
    return false;
  }

  try {
    await t.sendMail({
      from: GMAIL_USER,
      to: ADMIN_EMAIL,
      subject,
      text: body,
    });
    return true;
  } catch (err) {
    console.error('[email] 발송 실패:', err);
    return false;
  }
}
