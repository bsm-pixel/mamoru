-- 011: 이카운트 ERP 연동 제거
-- 컬럼은 유지 (기존 데이터 보존), 기본값만 제거하여 신규 레코드에 쓰지 않음

-- offline_sales: ecount_sync_status 기본값 제거
ALTER TABLE offline_sales ALTER COLUMN ecount_sync_status DROP DEFAULT;

-- 이카운트 인덱스 제거 (존재하는 경우)
DROP INDEX IF EXISTS idx_offline_sales_ecount;
DROP INDEX IF EXISTS idx_contracts_ecount;
