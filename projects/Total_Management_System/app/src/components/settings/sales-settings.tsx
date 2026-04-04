'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Plus, X, Upload } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';
import toast from 'react-hot-toast';
import type { TabProps } from '@/app/(dashboard)/settings/page';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

export default function SalesSettings({ settings, onSave, saving }: TabProps) {
  const supabase = createClient();
  const [businessInfo, setBusinessInfo] = useState({
    company: '', registration_number: '', representative: '',
    address: '', phone: '', business_type: '', business_item: '',
  });
  const [logoUrl, setLogoUrl] = useState('');
  const [receiptFooter, setReceiptFooter] = useState('');
  const [paymentMethods, setPaymentMethods] = useState<string[]>(['card', 'cash', 'transfer', 'mixed']);
  const [newMethod, setNewMethod] = useState('');
  const [channels, setChannels] = useState<string[]>(['offline', 'online', 'talk']);
  const [newChannel, setNewChannel] = useState('');
  const [dealerDiscount, setDealerDiscount] = useState(0);
  const [academyDiscount, setAcademyDiscount] = useState(0);
  const [blockZeroStock, setBlockZeroStock] = useState(false);
  const [defaultVat, setDefaultVat] = useState(true);
  const [autoTaxInvoice, setAutoTaxInvoice] = useState(false);
  const [uploading, setUploading] = useState(false);

  useEffect(() => {
    setBusinessInfo(parse(settings['business.info'], businessInfo));
    setLogoUrl(parse(settings['business.logo_url'], ''));
    setReceiptFooter(parse(settings['sales.receipt_footer'], ''));
    setPaymentMethods(parse(settings['sales.payment_methods'], ['card', 'cash', 'transfer', 'mixed']));
    setChannels(parse(settings['sales.channels'], ['offline', 'online', 'talk']));
    setDealerDiscount(parse(settings['sales.dealer_discount'], 0));
    setAcademyDiscount(parse(settings['sales.academy_discount'], 0));
    setBlockZeroStock(parse(settings['sales.block_zero_stock'], false));
    setDefaultVat(parse(settings['sales.default_vat_included'], true));
    setAutoTaxInvoice(parse(settings['sales.auto_tax_invoice'], false));
  }, [settings]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSave = () => {
    onSave([
      { key: 'business.info', value: businessInfo },
      { key: 'business.logo_url', value: logoUrl },
      { key: 'sales.receipt_footer', value: receiptFooter },
      { key: 'sales.payment_methods', value: paymentMethods },
      { key: 'sales.channels', value: channels },
      { key: 'sales.dealer_discount', value: dealerDiscount },
      { key: 'sales.academy_discount', value: academyDiscount },
      { key: 'sales.block_zero_stock', value: blockZeroStock },
      { key: 'sales.default_vat_included', value: defaultVat },
      { key: 'sales.auto_tax_invoice', value: autoTaxInvoice },
    ]);
  };

  const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const ext = file.name.split('.').pop() || 'png';
      const filePath = `settings/logo_${Date.now()}.${ext}`;
      const { error } = await supabase.storage.from('repair-photos').upload(filePath, file, { contentType: file.type, upsert: true });
      if (error) throw error;
      const { data: urlData } = supabase.storage.from('repair-photos').getPublicUrl(filePath);
      setLogoUrl(urlData.publicUrl);
      toast.success('로고 업로드 완료');
    } catch (err) {
      toast.error('로고 업로드 실패: ' + String(err));
    } finally {
      setUploading(false);
      e.target.value = '';
    }
  };

  const METHOD_LABELS: Record<string, string> = { card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합' };
  const CHANNEL_LABELS: Record<string, string> = { offline: '오프라인', online: '온라인', talk: '톡상담' };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">판매 관리 설정</h2>

      {/* 1. 거래명세서 발행인 정보 */}
      <Field label="거래명세서 / 세금계산서 발행인 정보" desc="모든 공식 문서에 자동 삽입됩니다. (회계 설정과 공유)">
        <div className="grid grid-cols-2 gap-2">
          <input placeholder="상호" value={businessInfo.company} onChange={(e) => setBusinessInfo({ ...businessInfo, company: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="사업자번호" value={businessInfo.registration_number} onChange={(e) => setBusinessInfo({ ...businessInfo, registration_number: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="대표자" value={businessInfo.representative} onChange={(e) => setBusinessInfo({ ...businessInfo, representative: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="연락처" value={businessInfo.phone} onChange={(e) => setBusinessInfo({ ...businessInfo, phone: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="주소" value={businessInfo.address} onChange={(e) => setBusinessInfo({ ...businessInfo, address: e.target.value })}
            className="col-span-2 h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="업태" value={businessInfo.business_type} onChange={(e) => setBusinessInfo({ ...businessInfo, business_type: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="종목" value={businessInfo.business_item} onChange={(e) => setBusinessInfo({ ...businessInfo, business_item: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>
      </Field>

      {/* 2. 로고 */}
      <Field label="거래명세서 로고 이미지" desc="A4 출력 좌측 상단에 표시됩니다.">
        <div className="flex items-center gap-3">
          {logoUrl && <img src={logoUrl} alt="로고" className="h-12 object-contain rounded border" />}
          <label className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg bg-neutral-100 text-sm cursor-pointer hover:bg-neutral-200">
            <Upload size={14} />
            {uploading ? '업로드 중...' : '로고 업로드'}
            <input type="file" accept="image/*" onChange={handleLogoUpload} disabled={uploading} className="hidden" />
          </label>
          {logoUrl && <button onClick={() => setLogoUrl('')} className="text-xs text-red-500">삭제</button>}
        </div>
      </Field>

      {/* 3. 하단 문구 */}
      <Field label="거래명세서 하단 문구" desc="교환/반품 안내 등.">
        <textarea value={receiptFooter} onChange={(e) => setReceiptFooter(e.target.value)} rows={3}
          className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm" placeholder="교환/반품은 수령 후 7일 이내..." />
      </Field>

      {/* 4. 결제 방법 */}
      <Field label="결제 방법 목록" desc="판매 입력 시 선택 가능한 결제 수단.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {paymentMethods.map((m) => (
            <span key={m} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
              {METHOD_LABELS[m] || m}
              <button onClick={() => setPaymentMethods(paymentMethods.filter((x) => x !== m))} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
            </span>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newMethod} onChange={(e) => setNewMethod(e.target.value)} placeholder="새 결제수단"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            onKeyDown={(e) => { if (e.key === 'Enter' && newMethod.trim()) { setPaymentMethods([...paymentMethods, newMethod.trim()]); setNewMethod(''); } }} />
          <button onClick={() => { if (newMethod.trim()) { setPaymentMethods([...paymentMethods, newMethod.trim()]); setNewMethod(''); } }}
            className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 5. 판매 채널 */}
      <Field label="판매 채널 목록" desc="판매 입력 시 선택 가능한 채널.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {channels.map((c) => (
            <span key={c} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
              {CHANNEL_LABELS[c] || c}
              <button onClick={() => setChannels(channels.filter((x) => x !== c))} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
            </span>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newChannel} onChange={(e) => setNewChannel(e.target.value)} placeholder="새 채널"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            onKeyDown={(e) => { if (e.key === 'Enter' && newChannel.trim()) { setChannels([...channels, newChannel.trim()]); setNewChannel(''); } }} />
          <button onClick={() => { if (newChannel.trim()) { setChannels([...channels, newChannel.trim()]); setNewChannel(''); } }}
            className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 6. 할인율 */}
      <Field label="B2B 기본 할인율" desc="딜러/아카데미 고객에게 자동 적용되는 할인.">
        <div className="flex gap-4">
          <div className="flex items-center gap-1">
            <span className="text-sm">딜러</span>
            <input type="number" value={dealerDiscount} onChange={(e) => setDealerDiscount(Number(e.target.value))}
              className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" min={0} max={100} />
            <span className="text-sm">%</span>
          </div>
          <div className="flex items-center gap-1">
            <span className="text-sm">아카데미</span>
            <input type="number" value={academyDiscount} onChange={(e) => setAcademyDiscount(Number(e.target.value))}
              className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" min={0} max={100} />
            <span className="text-sm">%</span>
          </div>
        </div>
      </Field>

      {/* 7. 재고 부족 시 판매 차단 */}
      <Field label="재고 부족 시 판매 차단" desc="재고 0인 상품 선택 시 판매 불가 처리.">
        <Toggle checked={blockZeroStock} onChange={setBlockZeroStock} />
      </Field>

      {/* 9. 부가세 기본 */}
      <Field label="기본 부가세 포함 여부" desc="판매 입력 시 VAT 포함이 기본 선택.">
        <Toggle checked={defaultVat} onChange={setDefaultVat} />
      </Field>

      {/* 16. 자동 세금계산서 */}
      <Field label="부가세 포함 거래 자동 세금계산서" desc="부가세 포함 판매 시 세금계산서 대장에 자동 등록.">
        <Toggle checked={autoTaxInvoice} onChange={setAutoTaxInvoice} />
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}>
          <Save size={14} />
          {saving ? '저장 중...' : '저장'}
        </Button>
      </div>
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>
      {desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}
      {children}
    </div>
  );
}

function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <button onClick={() => onChange(!checked)}
      className={`relative w-11 h-6 rounded-full transition ${checked ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
      <span className={`absolute top-0.5 left-0.5 w-5 h-5 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-5' : ''}`} />
    </button>
  );
}
