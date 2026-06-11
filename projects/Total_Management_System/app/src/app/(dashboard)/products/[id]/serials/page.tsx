'use client';

import { use } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { useProducts } from '@/hooks/use-sales';
import { SerialManagePanel } from '@/components/serials/serial-manage-panel';
import { ArrowLeft, Package } from 'lucide-react';

export default function SerialsPage({ params }: { params: Promise<{ id: string }> }) {
  const { id: productId } = use(params);
  const router = useRouter();
  const { data: products = [] } = useProducts();
  const product = products.find((p) => p.id === productId);

  return (
    <>
      <Topbar title={product ? `${product.name} — 시리얼 관리` : '시리얼 관리'} />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="flex items-center gap-3">
          <Button variant="ghost" size="sm" onClick={() => router.push('/products')}>
            <ArrowLeft size={14} />
          </Button>
          {product && (
            <div className="flex items-center gap-2">
              <Package size={16} className="text-stone-900" />
              <span className="text-sm font-bold">{product.name}</span>
              <span className="text-xs text-neutral-500">({product.sku})</span>
            </div>
          )}
        </div>

        <SerialManagePanel productId={productId} productName={product?.name} />
      </div>
    </>
  );
}
