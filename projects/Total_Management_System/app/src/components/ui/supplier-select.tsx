'use client';

import { useState, useRef, useEffect } from 'react';
import { useCustomers } from '@/hooks/use-customers';
import { ChevronDown, X } from 'lucide-react';

interface SupplierSelectProps {
  value: string;          // supplier_id
  displayName?: string;   // 초기 표시용 이름
  onChange: (id: string, name: string) => void;
  placeholder?: string;
}

export function SupplierSelect({ value, displayName, onChange, placeholder = '매입처 선택' }: SupplierSelectProps) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState('');
  const ref = useRef<HTMLDivElement>(null);

  const { data } = useCustomers({ type: 'supplier', limit: 50 });
  const suppliers = data?.customers || [];

  const filtered = search
    ? suppliers.filter((s) =>
        s.name.toLowerCase().includes(search.toLowerCase()) ||
        (s.company_name && s.company_name.toLowerCase().includes(search.toLowerCase()))
      )
    : suppliers;

  const selectedName = displayName || suppliers.find((s) => s.id === value)?.name || '';

  // 외부 클릭 시 닫기
  useEffect(() => {
    function handleClick(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, []);

  function handleSelect(id: string, name: string) {
    onChange(id, name);
    setSearch('');
    setOpen(false);
  }

  function handleClear() {
    onChange('', '');
    setSearch('');
  }

  return (
    <div ref={ref} className="relative">
      <div
        onClick={() => setOpen(!open)}
        className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm flex items-center cursor-pointer hover:border-neutral-300 focus-within:ring-2 focus-within:ring-terracotta/40"
      >
        {value && selectedName ? (
          <div className="flex items-center justify-between w-full">
            <span className="truncate">{selectedName}</span>
            <button
              onClick={(e) => { e.stopPropagation(); handleClear(); }}
              className="w-5 h-5 rounded flex items-center justify-center hover:bg-neutral-200"
            >
              <X size={12} />
            </button>
          </div>
        ) : (
          <div className="flex items-center justify-between w-full text-neutral-400">
            <span>{placeholder}</span>
            <ChevronDown size={14} />
          </div>
        )}
      </div>

      {open && (
        <div className="absolute z-50 mt-1 w-full bg-white border border-neutral-200 rounded-lg shadow-lg max-h-60 overflow-hidden">
          <div className="p-2 border-b border-neutral-100">
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="매입처 검색..."
              autoFocus
              className="w-full h-8 px-2 rounded border border-neutral-200 bg-neutral-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-terracotta/40"
            />
          </div>
          <div className="overflow-y-auto max-h-44">
            {filtered.length === 0 ? (
              <p className="text-xs text-neutral-400 text-center py-3">매입처 없음</p>
            ) : (
              filtered.map((s) => (
                <button
                  key={s.id}
                  onClick={() => handleSelect(s.id, s.company_name || s.name)}
                  className={`w-full text-left px-3 py-2 text-sm hover:bg-neutral-50 transition ${
                    s.id === value ? 'bg-terracotta/5 text-terracotta font-medium' : ''
                  }`}
                >
                  <p className="truncate">{s.company_name || s.name}</p>
                  {s.company_name && s.name !== s.company_name && (
                    <p className="text-xs text-neutral-400 truncate">{s.name}</p>
                  )}
                </button>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
}
