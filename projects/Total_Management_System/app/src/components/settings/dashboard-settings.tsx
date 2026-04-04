'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save } from 'lucide-react';
import type { TabProps } from '@/app/(dashboard)/settings/page';

function parse<T>(raw: unknown, fallback: T): T {
  if (raw === undefined || raw === null) return fallback;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

export default function DashboardSettings({ settings, onSave, saving }: TabProps) {
  const [monthlyGoal, setMonthlyGoal] = useState(0);
  const [lowStock, setLowStock] = useState(3);
  const [staleDays, setStaleDays] = useState(3);
  const [kpiGreen, setKpiGreen] = useState(80);
  const [kpiYellow, setKpiYellow] = useState(50);
  const [outstandingWarning, setOutstandingWarning] = useState(0);
  const [purchaseLimit, setPurchaseLimit] = useState(0);
  const [cardVisibility, setCardVisibility] = useState<Record<string, boolean>>({
    sales: true, repairs: true, orders: true, consultations: true, outstanding: true, lowStock: true,
  });
  const [cardOrder, setCardOrder] = useState<string[]>([
    'sales', 'repairs', 'orders', 'consultations', 'outstanding', 'lowStock',
  ]);

  useEffect(() => {
    setMonthlyGoal(parse(settings['dashboard.monthly_goal'], 0));
    setLowStock(parse(settings['dashboard.low_stock_threshold'], 3));
    setStaleDays(parse(settings['dashboard.repair_stale_days'], 3));
    setKpiGreen(parse(settings['dashboard.kpi_green'], 80));
    setKpiYellow(parse(settings['dashboard.kpi_yellow'], 50));
    setOutstandingWarning(parse(settings['dashboard.outstanding_warning'], 0));
    setPurchaseLimit(parse(settings['dashboard.monthly_purchase_limit'], 0));
    setCardVisibility(parse(settings['dashboard.card_visibility'], {
      sales: true, repairs: true, orders: true, consultations: true, outstanding: true, lowStock: true,
    }));
    setCardOrder(parse(settings['dashboard.card_order'], [
      'sales', 'repairs', 'orders', 'consultations', 'outstanding', 'lowStock',
    ]));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'dashboard.monthly_goal', value: monthlyGoal },
      { key: 'dashboard.low_stock_threshold', value: lowStock },
      { key: 'dashboard.repair_stale_days', value: staleDays },
      { key: 'dashboard.kpi_green', value: kpiGreen },
      { key: 'dashboard.kpi_yellow', value: kpiYellow },
      { key: 'dashboard.outstanding_warning', value: outstandingWarning },
      { key: 'dashboard.monthly_purchase_limit', value: purchaseLimit },
      { key: 'dashboard.card_visibility', value: cardVisibility },
      { key: 'dashboard.card_order', value: cardOrder },
    ]);
  };

  const CARD_LABELS: Record<string, string> = {
    sales: '매출', repairs: '복원수리', orders: '주문', consultations: '상담', outstanding: '미수금', lowStock: '저재고',
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">대시보드 설정</h2>

      {/* 1. 월 매출 목표 */}
      <Field label="월 매출 목표 금액" desc="KPI 게이지 기준값. 모든 기기에서 동일하게 표시됩니다.">
        <div className="flex items-center gap-2">
          <input type="number" value={monthlyGoal} onChange={(e) => setMonthlyGoal(Number(e.target.value))}
            className="w-40 h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <span className="text-sm text-neutral-500">원</span>
        </div>
      </Field>

      {/* 2. 저재고 알림 기준 */}
      <Field label="저재고 알림 기준 수량" desc="이 수량 이하인 상품이 대시보드 저재고 카드에 표시됩니다.">
        <input type="number" value={lowStock} onChange={(e) => setLowStock(Number(e.target.value))}
          className="w-24 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
      </Field>

      {/* 3. 복원수리 체류 경고 일수 */}
      <Field label="복원수리 체류 경고 일수" desc="상태 변경 없이 이 일수를 넘으면 '정체' 경고가 표시됩니다.">
        <div className="flex items-center gap-2">
          <input type="number" value={staleDays} onChange={(e) => setStaleDays(Number(e.target.value))}
            className="w-24 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={1} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      {/* 4+8. 대시보드 카드 배치 — 실제 그리드 프리뷰 */}
      <Field label="대시보드 카드 배치" desc="실제 대시보드와 동일한 2열 그리드. 클릭으로 숨김/표시, 드래그하여 순서 변경.">
        <div className="grid grid-cols-2 gap-2 p-3 bg-neutral-50 rounded-lg border border-neutral-200">
          {cardOrder.map((key, idx) => {
            const visible = cardVisibility[key] !== false;
            return (
              <div
                key={key}
                className={`relative rounded-lg border-2 p-3 text-center transition cursor-move select-none
                  ${visible
                    ? 'bg-white border-neutral-300 shadow-sm'
                    : 'bg-neutral-100 border-dashed border-neutral-200 opacity-40'
                  }`}
                draggable
                onDragStart={(e) => e.dataTransfer.setData('idx', String(idx))}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  const fromIdx = Number(e.dataTransfer.getData('idx'));
                  if (fromIdx === idx) return;
                  const next = [...cardOrder];
                  const [moved] = next.splice(fromIdx, 1);
                  next.splice(idx, 0, moved);
                  setCardOrder(next);
                }}
              >
                <div className="text-sm font-semibold">{CARD_LABELS[key] || key}</div>
                <div className="text-[10px] text-neutral-400 mt-0.5">
                  {visible ? `${idx + 1}번 위치` : '숨김'}
                </div>
                <button
                  onClick={() => setCardVisibility({ ...cardVisibility, [key]: !visible })}
                  className={`absolute top-1 right-1 w-5 h-5 rounded-full text-[10px] font-bold flex items-center justify-center
                    ${visible ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-500'}`}
                >
                  {visible ? '✓' : '✕'}
                </button>
              </div>
            );
          })}
        </div>
        <p className="text-[10px] text-neutral-400 mt-1.5">드래그로 순서 변경 · 우측 상단 버튼으로 숨김/표시</p>
      </Field>

      {/* 5. KPI 색상 임계치 */}
      <Field label="KPI 달성률 색상 기준" desc="게이지 색상이 바뀌는 기준 퍼센트입니다.">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1">
            <span className="w-3 h-3 rounded-full bg-green-500" />
            <input type="number" value={kpiGreen} onChange={(e) => setKpiGreen(Number(e.target.value))}
              className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" />
            <span className="text-xs text-neutral-500">% 이상</span>
          </div>
          <div className="flex items-center gap-1">
            <span className="w-3 h-3 rounded-full bg-yellow-500" />
            <input type="number" value={kpiYellow} onChange={(e) => setKpiYellow(Number(e.target.value))}
              className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" />
            <span className="text-xs text-neutral-500">% 이상</span>
          </div>
          <div className="flex items-center gap-1">
            <span className="w-3 h-3 rounded-full bg-red-500" />
            <span className="text-xs text-neutral-500">미만</span>
          </div>
        </div>
      </Field>

      {/* 6. 미수금 경고 기준 금액 */}
      <Field label="미수금 경고 기준 금액" desc="이 금액 이상 미수금만 대시보드에 강조 표시됩니다. 0이면 전부 표시.">
        <div className="flex items-center gap-2">
          <input type="number" value={outstandingWarning} onChange={(e) => setOutstandingWarning(Number(e.target.value))}
            className="w-40 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} step={10000} />
          <span className="text-sm text-neutral-500">원</span>
        </div>
      </Field>

      {/* 7. 월 목표 매입 한도 */}
      <Field label="월 목표 매입 한도" desc="매입 페이지에서 잔여 예산을 표시합니다. 0이면 비활성.">
        <div className="flex items-center gap-2">
          <input type="number" value={purchaseLimit} onChange={(e) => setPurchaseLimit(Number(e.target.value))}
            className="w-40 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} step={100000} />
          <span className="text-sm text-neutral-500">원</span>
        </div>
      </Field>

      {/* 8. 카드 순서 — 위 그리드 프리뷰에 통합됨 */}

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}>
          <Save size={14} />
          {saving ? '저장 중...' : '저장'}
        </Button>
      </div>
    </div>
  );
}

/* ── 공통 필드 래퍼 ── */
function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>
      {desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}
      {children}
    </div>
  );
}
