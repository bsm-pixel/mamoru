import { describe, it, expect } from 'vitest';
import { pickPrinter } from '@/lib/label/browser-print';

// 라벨 프린터 고르기(2026-11-20) — 공유기가 프린터 주소를 바꿔도, 기본값이 윈도우 드라이버로 잡혀도 출력되게
const driver = { uid: 'ZDesigner ZT231-203dpi ZPL', connection: 'driver', manufacturer: 'Zebra Technologies' };
const net = (ip: string) => ({ uid: ip, connection: 'network', manufacturer: 'Zebra Technologies' });

describe('라벨 프린터 고르기', () => {
  it('기본 프린터가 지금도 찾아지면 그대로 쓴다', () => {
    expect(pickPrinter(net('192.168.219.41'), [driver, net('192.168.219.41')])?.uid).toBe('192.168.219.41');
  });
  it('기본 프린터 주소가 옛것이면 지금 찾아진 주소로 보낸다', () => {
    expect(pickPrinter(net('192.168.219.44'), [driver, net('192.168.219.41')])?.uid).toBe('192.168.219.41');
  });
  it('기본 프린터가 윈도우 드라이버여도 주소로 찾아진 프린터를 먼저 쓴다', () => {
    expect(pickPrinter(driver, [driver, net('192.168.219.41')])?.uid).toBe('192.168.219.41');
  });
  it('찾아진 프린터가 없으면 기본 프린터를 쓴다(전과 같음)', () => {
    expect(pickPrinter(driver, [driver])?.uid).toBe(driver.uid);
    expect(pickPrinter(net('192.168.219.44'), null)?.uid).toBe('192.168.219.44');
    expect(pickPrinter(null, [])).toBeNull();
  });
});
