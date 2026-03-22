-- 026: 판매 취소/채널 확장
-- 1) 취소 관련 컬럼
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS cancelled_at timestamptz DEFAULT NULL;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS cancelled_reason text DEFAULT NULL;
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS cancelled_by uuid REFERENCES profiles(id) DEFAULT NULL;

-- 2) 채널 컬럼 (기존 데이터 기본값 'offline')
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS sale_channel text NOT NULL DEFAULT 'offline';
-- sale_channel: 'offline' | 'online' | 'talk'

-- 3) 인덱스
CREATE INDEX IF NOT EXISTS idx_offline_sales_channel ON offline_sales(sale_channel);
CREATE INDEX IF NOT EXISTS idx_offline_sales_cancelled ON offline_sales(cancelled_at) WHERE cancelled_at IS NOT NULL;
