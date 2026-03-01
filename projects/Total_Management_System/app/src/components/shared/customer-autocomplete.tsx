'use client';

import { useState, useRef, useEffect, useCallback } from 'react';
import { Search, X, Plus, User, Loader2, RefreshCw } from 'lucide-react';
import { useCustomerSearch, useCreateCustomer, useSyncEcountCustomers, type CustomerResult } from '@/hooks/use-customers';

export interface SelectedCustomer {
  id: string;
  name: string;
  phone: string | null;
  email: string | null;
  address_road: string | null;
  address_detail: string | null;
  postcode: string | null;
  ecount_customer_code: string | null;
}

interface Props {
  selectedCustomer: SelectedCustomer | null;
  onSelect: (customer: SelectedCustomer) => void;
  onClear: () => void;
  /** 계약서용: email/address 필드도 신규등록 폼에 표시 */
  showExtendedFields?: boolean;
}

const SOURCE_LABEL: Record<string, string> = {
  imweb: '온라인',
  consultation: '상담',
  as: '복원수리',
  manual: '수동',
};

export function CustomerAutocomplete({ selectedCustomer, onSelect, onClear, showExtendedFields }: Props) {
  const [query, setQuery] = useState('');
  const [debouncedQuery, setDebouncedQuery] = useState('');
  const [showDropdown, setShowDropdown] = useState(false);
  const [showNewForm, setShowNewForm] = useState(false);

  // 신규 등록 폼 state
  const [newName, setNewName] = useState('');
  const [newPhone, setNewPhone] = useState('');
  const [newEmail, setNewEmail] = useState('');
  const [newAddress, setNewAddress] = useState('');

  const containerRef = useRef<HTMLDivElement>(null);
  const timerRef = useRef<NodeJS.Timeout>(undefined);

  const { data: results = [], isLoading } = useCustomerSearch(debouncedQuery);
  const createCustomer = useCreateCustomer();
  const syncEcount = useSyncEcountCustomers();

  // 디바운스 검색
  const handleQueryChange = useCallback((value: string) => {
    setQuery(value);
    if (timerRef.current) clearTimeout(timerRef.current);
    if (value.length >= 2) {
      timerRef.current = setTimeout(() => setDebouncedQuery(value), 300);
      setShowDropdown(true);
      setShowNewForm(false);
    } else {
      setDebouncedQuery('');
      setShowDropdown(false);
    }
  }, []);

  // 외부 클릭 시 드롭다운 닫기
  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setShowDropdown(false);
        setShowNewForm(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // 고객 선택
  function handleSelect(customer: CustomerResult) {
    onSelect({
      id: customer.id,
      name: customer.name,
      phone: customer.phone,
      email: customer.email,
      address_road: customer.address_road,
      address_detail: customer.address_detail,
      postcode: customer.postcode,
      ecount_customer_code: customer.ecount_customer_code,
    });
    setQuery('');
    setDebouncedQuery('');
    setShowDropdown(false);
    setShowNewForm(false);
  }

  // 신규 등록 폼 열기
  function openNewForm() {
    setNewName(query); // 검색어를 이름에 프리필
    setNewPhone('');
    setNewEmail('');
    setNewAddress('');
    setShowNewForm(true);
    setShowDropdown(false);
  }

  // 신규 등록 제출
  async function handleCreateCustomer() {
    if (!newName.trim()) return;

    const result = await createCustomer.mutateAsync({
      name: newName.trim(),
      phone: newPhone.trim() || undefined,
      email: newEmail.trim() || undefined,
      address: newAddress.trim() || undefined,
    });

    onSelect({
      id: result.customer.id,
      name: result.customer.name,
      phone: result.customer.phone,
      email: result.customer.email,
      address_road: result.customer.address_road,
      address_detail: result.customer.address_detail,
      postcode: result.customer.postcode,
      ecount_customer_code: result.customer.ecount_customer_code,
    });
    setQuery('');
    setDebouncedQuery('');
    setShowNewForm(false);
  }

  // 선택 완료 상태
  if (selectedCustomer) {
    return (
      <div className="space-y-2">
        <div className="flex items-center gap-2 p-2.5 rounded-lg border border-terracotta/30 bg-terracotta/5">
          <User size={14} className="text-terracotta shrink-0" />
          <div className="flex-1 min-w-0">
            <span className="text-sm font-semibold text-indigo-black">{selectedCustomer.name}</span>
            {selectedCustomer.phone && (
              <span className="text-xs text-neutral-500 ml-2">{selectedCustomer.phone}</span>
            )}
            {selectedCustomer.ecount_customer_code && (
              <span className="text-[10px] text-terracotta ml-2 font-mono">
                {selectedCustomer.ecount_customer_code}
              </span>
            )}
          </div>
          <button
            type="button"
            onClick={onClear}
            className="w-5 h-5 rounded-full bg-neutral-200 flex items-center justify-center hover:bg-neutral-300"
          >
            <X size={10} />
          </button>
        </div>
        {/* 선택된 고객의 연락처 (readonly) */}
        {selectedCustomer.phone && (
          <div className="text-xs text-neutral-500 px-1">
            연락처: {selectedCustomer.phone}
          </div>
        )}
      </div>
    );
  }

  // 신규 등록 폼
  if (showNewForm) {
    return (
      <div ref={containerRef} className="space-y-2">
        <div className="p-3 rounded-lg border border-terracotta/30 bg-warm-ivory space-y-2">
          <p className="text-xs font-semibold text-terracotta">신규 고객 등록</p>
          <input
            type="text"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            placeholder="고객명 *"
            className="w-full h-8 px-3 rounded border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-terracotta/40"
          />
          <input
            type="tel"
            value={newPhone}
            onChange={(e) => setNewPhone(e.target.value)}
            placeholder="연락처 (010-0000-0000)"
            className="w-full h-8 px-3 rounded border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-terracotta/40"
          />
          {showExtendedFields && (
            <>
              <input
                type="email"
                value={newEmail}
                onChange={(e) => setNewEmail(e.target.value)}
                placeholder="이메일"
                className="w-full h-8 px-3 rounded border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-terracotta/40"
              />
              <input
                type="text"
                value={newAddress}
                onChange={(e) => setNewAddress(e.target.value)}
                placeholder="주소"
                className="w-full h-8 px-3 rounded border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-terracotta/40"
              />
            </>
          )}
          <div className="flex gap-2 pt-1">
            <button
              type="button"
              onClick={() => { setShowNewForm(false); setShowDropdown(true); }}
              className="flex-1 h-8 rounded text-xs font-semibold bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
            >
              취소
            </button>
            <button
              type="button"
              onClick={handleCreateCustomer}
              disabled={!newName.trim() || createCustomer.isPending}
              className="flex-1 h-8 rounded text-xs font-semibold bg-terracotta text-cream hover:bg-terracotta/90 disabled:opacity-50"
            >
              {createCustomer.isPending ? '등록중...' : '등록 + 선택'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // 검색 상태
  return (
    <div ref={containerRef} className="relative">
      <div className="flex gap-1.5">
        <div className="relative flex-1">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
          <input
            type="text"
            value={query}
            onChange={(e) => handleQueryChange(e.target.value)}
            onFocus={() => { if (debouncedQuery.length >= 2) setShowDropdown(true); }}
            placeholder="고객명 또는 전화번호 검색..."
            className="w-full h-9 pl-8 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        {/* 이카운트 거래처 동기화 */}
        <button
          type="button"
          onClick={() => syncEcount.mutate()}
          disabled={syncEcount.isPending}
          title="이카운트 거래처 동기화"
          className="h-9 px-2 rounded-lg border border-neutral-200 bg-warm-ivory hover:bg-neutral-100 text-neutral-500 hover:text-terracotta disabled:opacity-50 shrink-0"
        >
          <RefreshCw size={14} className={syncEcount.isPending ? 'animate-spin' : ''} />
        </button>
      </div>

      {/* 드롭다운 */}
      {showDropdown && debouncedQuery.length >= 2 && (
        <div className="absolute z-50 w-full mt-1 bg-white rounded-lg border border-neutral-200 shadow-lg max-h-64 overflow-y-auto">
          {isLoading ? (
            <div className="flex items-center justify-center py-4 text-neutral-400">
              <Loader2 size={16} className="animate-spin mr-2" />
              <span className="text-xs">검색중...</span>
            </div>
          ) : results.length > 0 ? (
            <>
              {results.map((c) => (
                <button
                  key={c.id}
                  type="button"
                  onClick={() => handleSelect(c)}
                  className="w-full px-3 py-2.5 text-left hover:bg-neutral-50 border-b border-neutral-50 last:border-b-0"
                >
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-semibold text-indigo-black">{c.name}</span>
                    {c.phone && (
                      <span className="text-xs text-neutral-500">{c.phone}</span>
                    )}
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-500 ml-auto">
                      {SOURCE_LABEL[c.source] || c.source}
                    </span>
                  </div>
                  {c.ecount_customer_code && (
                    <span className="text-[10px] text-terracotta/70 font-mono">
                      이카운트: {c.ecount_customer_code}
                    </span>
                  )}
                </button>
              ))}
              {/* 신규 등록 버튼 */}
              <button
                type="button"
                onClick={openNewForm}
                className="w-full px-3 py-2.5 text-left hover:bg-terracotta/5 border-t border-neutral-100 flex items-center gap-2"
              >
                <Plus size={14} className="text-terracotta" />
                <span className="text-sm text-terracotta font-semibold">
                  &quot;{query}&quot; 신규 등록
                </span>
              </button>
            </>
          ) : (
            <>
              <div className="px-3 py-3 text-center">
                <p className="text-xs text-neutral-400">등록된 고객이 없습니다</p>
              </div>
              <button
                type="button"
                onClick={openNewForm}
                className="w-full px-3 py-2.5 text-left hover:bg-terracotta/5 border-t border-neutral-100 flex items-center gap-2"
              >
                <Plus size={14} className="text-terracotta" />
                <span className="text-sm text-terracotta font-semibold">
                  &quot;{query}&quot; 신규 등록
                </span>
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
}
