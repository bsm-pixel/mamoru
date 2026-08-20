/** Supabase Database 타입 — 수동 정의 (Phase 1 + Phase 2-1) */

export type OrderStatus =
  | 'pay_wait'
  | 'pay_done'
  | 'preparing'
  | 'shipping'
  | 'delivered'
  | 'confirmed'
  | 'cancel_pending'
  | 'cancelled'
  | 'refund_request'
  | 'refunded';

export type UserRole = 'owner' | 'staff' | 'director';

// Phase 2-1: 상담/딜러 ENUM
export type ConsultationStatus =
  | 'pending_admin'
  | 'suggested'
  | 'assigned'
  | 'confirmed'
  | 'cancelled'
  | 'reschedule_requested'
  | 'change_requested'  // Phase 1.6: 고객 일정변경 요청 (출장)
  | 'on_hold'       // Phase 2-2: 보류
  | 'in_progress'   // Phase 2-2: 진행중 (톡상담)
  | 'completed';    // Phase 2-2: 처리완료

export type ConsultationType = 'store_visit' | 'field_request' | 'talk_consult';

export type DealerStatus = 'active' | 'inactive';

// Phase 7: 복원수리 ENUM
export type RepairStatus =
  | 'intake'
  | 'pickup_scheduled'
  | 'picked_up'
  | 'inspecting'
  | 'cost_notified'
  | 'payment_confirmed'
  | 'repairing'
  | 'ready_to_ship'
  | 'shipped'
  | 'delivered'
  | 'completed'
  | 'cancelled';

// R5: 오프라인 판매 ENUM
export type PaymentMethod = 'card' | 'cash' | 'transfer' | 'mixed';
export type PaymentStatus = 'paid' | 'unpaid' | 'partial';
export type SaleChannel = 'offline' | 'online' | 'talk';
// EcountSyncStatus 제거됨 (이카운트 연동 폐기)

// R6: 전자 계약서 ENUM
export type ContractStatus = 'draft' | 'signed' | 'sent' | 'completed' | 'cancelled';

// R7: 시리얼넘버 ENUM
export type SerialStatus = 'in_stock' | 'reserved' | 'sold' | 'returned' | 'defective';
export type WarehouseZone = 'raw' | 'ready' | 'display';

/** 112: 로케이션의 용도 구역. WarehouseZone(시리얼의 존)과 개념은 같으나
 *  raw 존의 물리적 자리를 다루므로 'storage' 라는 이름을 쓴다. */
export type WarehouseZoneType = 'storage' | 'ready' | 'display';

export interface Database {
  public: {
    Tables: {
      profiles: {
        Row: {
          id: string;
          display_name: string | null;
          role: UserRole;
          phone: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['profiles']['Row'], 'created_at' | 'updated_at'>;
        Update: Partial<Database['public']['Tables']['profiles']['Insert']>;
      };
      customers: {
        Row: {
          id: string;
          name: string;
          phone: string | null;
          phone_normalized: string | null;
          email: string | null;
          postcode: string | null;
          address_road: string | null;
          address_detail: string | null;
          source: 'imweb' | 'consultation' | 'as' | 'manual';
          customer_type: 'retail' | 'online' | 'dealer' | 'academy' | 'supplier';
          company_name: string | null;
          activity_name: string | null;          // 102: 활동명(매장 사용 이름, 예 하은)
          position: string | null;               // 102: 직급(원장/디자이너 등)
          memo: string | null;
          tags: string[] | null;                 // 060: 고객 태그
          outstanding_balance: number;
          default_repair_price: number | null;  // 079: 거래처별 복원수리 기본 단가(자루당, 원)
          total_orders: number;
          total_spent: number;
          ecount_customer_code: string | null;  // legacy
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['customers']['Row'], 'id' | 'created_at' | 'updated_at' | 'phone_normalized' | 'total_orders' | 'total_spent' | 'ecount_customer_code' | 'outstanding_balance' | 'default_repair_price' | 'activity_name' | 'position' | 'tags'>;
        Update: Partial<Database['public']['Tables']['customers']['Insert']>;
      };
      orders: {
        Row: {
          id: string;
          imweb_order_no: string;
          imweb_order_id: string | null;
          customer_id: string | null;
          orderer_name: string;
          orderer_phone: string | null;
          orderer_email: string | null;
          recipient_name: string;
          recipient_phone: string | null;
          recipient_postcode: string | null;
          recipient_address: string | null;
          recipient_address_detail: string | null;
          recipient_memo: string | null;
          total_price: number;
          delivery_fee: number;
          discount_amount: number;
          paid_amount: number;
          payment_method: string | null;
          paid_at: string | null;
          courier_code: string | null;
          courier_name: string | null;
          invoice_number: string | null;
          shipped_at: string | null;
          delivered_at: string | null;
          review_requested_at: string | null; // 011: 리뷰 요청 발송 시각
          is_pickup?: boolean | null;          // 126: 직접수령(대면 픽업) 마커
          status: OrderStatus;
          imweb_raw: Record<string, unknown> | null;
          synced_at: string;
          ordered_at: string;
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['orders']['Row'], 'id' | 'created_at' | 'updated_at'>;
        Update: Partial<Database['public']['Tables']['orders']['Insert']>;
      };
      order_items: {
        Row: {
          id: string;
          order_id: string;
          product_id: string | null;
          imweb_product_no: string | null;
          product_name: string;
          option_text: string | null;
          quantity: number;
          unit_price: number;
          total_price: number;
        };
        Insert: Omit<Database['public']['Tables']['order_items']['Row'], 'id'>;
        Update: Partial<Database['public']['Tables']['order_items']['Insert']>;
      };
      products: {
        Row: {
          id: string;
          sku: string;
          name: string;
          category: string;
          price: number;
          price_dealer: number;
          price_academy: number;
          price_purchase: number;
          supplier_id: string | null;
          description: string | null;
          imweb_product_no: string | null;
          image_url: string | null;
          tags: Record<string, unknown> | null;
          stock_quantity: number;
          raw_stock: number;
          barcode: string | null;
          product_group: string | null;
          purchase_name: string | null;
          price_groups: Record<string, { price?: number | null; display_name?: string | null }> | null;
          is_active: boolean;
          // 112: 정위치(보관 자리). NULL=미지정. 재고 수량과 무관한 위치 참조 (2026-07-18)
          location_id: string | null;
          created_at: string;
          updated_at: string;
        };
        // location_id 는 Omit 후 선택 필드로 다시 붙인다 — 기존 제품 생성 코드가 안 깨지도록(위치는 나중에 배정)
        Insert: Omit<Database['public']['Tables']['products']['Row'], 'id' | 'created_at' | 'updated_at' | 'stock_quantity' | 'is_active' | 'price_dealer' | 'price_academy' | 'price_purchase' | 'price_groups' | 'location_id'>
          & { location_id?: string | null };
        Update: Partial<Database['public']['Tables']['products']['Insert']> & { price_dealer?: number; price_academy?: number; price_purchase?: number; price_groups?: Record<string, { price?: number | null; display_name?: string | null }> };
      };
      /** 113: 렉 정보 — 렉을 N열 그리드로 보고 단마다 쓰는 칸 수를 다르게 (2026-07-18) */
      warehouse_racks: {
        Row: {
          id: string;
          rack_no: number;
          label: string | null;
          columns: number;          // 렉 전체 열 수
          sort_order: number;
          memo: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['warehouse_racks']['Row'], 'id' | 'created_at' | 'updated_at' | 'sort_order' | 'columns'>
          & { sort_order?: number; columns?: number };
        Update: Partial<Database['public']['Tables']['warehouse_racks']['Insert']>;
      };
      /** 112: 창고 로케이션(정위치) — 렉·단·칸 물리적 자리 (2026-07-18) */
      warehouse_locations: {
        Row: {
          id: string;
          code: string;              // 'R01-2-A'
          label: string | null;      // '1번렉 중단 A칸'
          rack_no: number;
          level_no: number;          // 1 = 상단 (위→아래)
          bin_no: number | null;     // 열 (null = 칸 없이 선반 통째)
          bin_row: number | null;    // 114: 행 (수납함이면 2 이상, null = 선반)
          zone_type: WarehouseZoneType;
          sort_order: number;
          is_active: boolean;
          memo: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['warehouse_locations']['Row'], 'id' | 'created_at' | 'updated_at' | 'sort_order' | 'is_active' | 'zone_type'>
          & { sort_order?: number; is_active?: boolean; zone_type?: WarehouseZoneType };
        Update: Partial<Database['public']['Tables']['warehouse_locations']['Insert']>;
      };
      sync_log: {
        Row: {
          id: string;
          sync_type: string;
          status: string;
          records_synced: number;
          error_message: string | null;
          started_at: string;
          completed_at: string | null;
        };
        Insert: Omit<Database['public']['Tables']['sync_log']['Row'], 'id'>;
        Update: Partial<Database['public']['Tables']['sync_log']['Insert']>;
      };
      // Phase 2-1: 상담/딜러 테이블
      dealers: {
        Row: {
          id: string;
          dealer_code: string;
          name: string;
          phone: string | null;
          regions: string[];
          calendar_id: string | null;
          status: DealerStatus;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          dealer_code: string;
          name: string;
          phone?: string | null;
          regions?: string[];
          calendar_id?: string | null;
          status?: DealerStatus;
        };
        Update: {
          dealer_code?: string;
          name?: string;
          phone?: string | null;
          regions?: string[];
          calendar_id?: string | null;
          status?: DealerStatus;
        };
      };
      consultations: {
        Row: {
          id: string;
          customer_id: string | null;
          name: string;
          phone: string;
          phone_normalized: string | null;
          consultation_type: ConsultationType;
          activity_name: string | null;     // 102: 활동명(매장 사용 이름)
          position: string | null;          // 102: 직급
          visit_date: string | null;
          visit_time: string | null;
          postcode: string | null;
          address_road: string | null;
          address_detail: string | null;
          address_sido: string | null;
          address_sigungu: string | null;
          address_region: string | null;
          status: ConsultationStatus;
          memo: string | null;
          admin_note: string | null;        // 108: 상담자(관리자) 전용 메모 — 고객 비노출
          unique_id: string;
          dealer_id: string | null;
          suggestions: Record<string, unknown> | null;
          gas_raw: Record<string, unknown> | null;
          hold_reason: string | null;       // Phase 2-2: 보류 사유
          latitude: number | null;          // Phase 2-2: 좌표
          longitude: number | null;
          review_promised_at: string | null;     // 067: 리뷰 약속 시점
          review_promised_type: 'purchase' | 'repair' | 'consult' | null; // 094: 리뷰 약속 유형
          review_promised_subtype: string | null; // 095: 약속 세부 유형 (direct_visit/pickup/store_visit/field_request/talk_consult) (자동 발송 분기)
          review_request_sent_at: string | null; // 067: 후기 요청 발송 시점
          review_submitted_at: string | null;    // 067: 리뷰 작성 완료 시점
          remind_24h_at: string | null;          // 071: 24h 리마인더 발송 시각
          remind_2h_at: string | null;           // 071: 2h 리마인더 발송 시각
          received_at: string;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          customer_id?: string | null;
          name: string;
          phone: string;
          consultation_type?: ConsultationType;
          activity_name?: string | null;
          position?: string | null;
          visit_date?: string | null;
          visit_time?: string | null;
          postcode?: string | null;
          address_road?: string | null;
          address_detail?: string | null;
          address_sido?: string | null;
          address_sigungu?: string | null;
          address_region?: string | null;
          status?: ConsultationStatus;
          memo?: string | null;
          unique_id: string;
          dealer_id?: string | null;
          suggestions?: Record<string, unknown> | null;
          gas_raw?: Record<string, unknown> | null;
          hold_reason?: string | null;
          latitude?: number | null;
          longitude?: number | null;
          received_at?: string;
        };
        Update: {
          customer_id?: string | null;
          name?: string;
          phone?: string;
          consultation_type?: ConsultationType;
          visit_date?: string | null;
          visit_time?: string | null;
          postcode?: string | null;
          address_road?: string | null;
          address_detail?: string | null;
          address_sido?: string | null;
          address_sigungu?: string | null;
          address_region?: string | null;
          status?: ConsultationStatus;
          memo?: string | null;
          admin_note?: string | null;       // 108: 상담자(관리자) 전용 메모
          dealer_id?: string | null;
          suggestions?: Record<string, unknown> | null;
          gas_raw?: Record<string, unknown> | null;
          hold_reason?: string | null;
          latitude?: number | null;
          longitude?: number | null;
          review_promised_at?: string | null;
          review_promised_type?: 'purchase' | 'repair' | 'consult' | null;
          review_promised_subtype?: string | null;
          review_request_sent_at?: string | null;
          review_submitted_at?: string | null;
          remind_24h_at?: string | null;
          remind_2h_at?: string | null;
          received_at?: string;
        };
      };
      // Phase 7: 복원수리 테이블
      repairs: {
        Row: {
          id: string;
          as_id: string;
          customer_id: string | null;
          name: string;
          phone: string;
          phone_normalized: string | null;
          proceed_type: string | null;
          postcode: string | null;
          address: string | null;
          address_detail: string | null;
          pickup_date: string | null;
          delivery_method: string | null;
          qty_mamoru: number;
          qty_other: number;
          memo: string | null;
          service_cost: number;
          shipping_fee: number;
          total_amount: number;
          status: RepairStatus;
          invoice_number: string | null;
          courier_name: string | null;
          shipped_at: string | null;
          shipped_source: 'manual' | 'alps_pickup' | null;  // 109: 출고 기록 주체 (집하 자동감지 구분)
          paid_at: string | null;
          delivered_at: string | null;
          confirmed_at: string | null;  // R1: 접수확인 시점
          packed_at: string | null;     // R1: 포장완료 시점
          inbound_at: string | null;    // 119: 입고 & 비용안내(cost_notified) 첫 전이 시점 = 목록 '입고일'
          payment_method: string | null; // 120: 결제수단 transfer/card/cash (직접방문 현장결제·입금확인 기록)
          manage_token: string | null;  // 121: 고객 셀프 관리(일정확인/변경/취소) 링크 토큰
          visit_remind_24h_sent_at: string | null;  // 125: 직접방문 리마인드 중복방지
          visit_remind_2h_sent_at: string | null;   // 125: 직접방문 리마인드 중복방지
          admin_note: string | null;
          gas_raw: Record<string, unknown> | null;
          review_promised_at: string | null;     // 067: 리뷰 약속 시점
          review_promised_type: 'purchase' | 'repair' | 'consult' | null; // 094: 리뷰 약속 유형
          review_promised_subtype: string | null; // 095: 약속 세부 유형 (direct_visit/pickup/store_visit/field_request/talk_consult)
          review_request_sent_at: string | null; // 067: 후기 요청 발송 시점
          review_submitted_at: string | null;    // 067: 리뷰 작성 완료 시점
          // 092: 직접방문(당일수리)
          visit_date: string | null;             // 'YYYY-MM-DD'
          visit_time: string | null;             // 'HH:MM' (30분 단위)
          visit_duration_min: number | null;     // 30 or 60 (서버 계산, 충돌 검사용)
          // 093: Google Calendar 단방향 동기화 (Phase 3-B) — consultations 052 와 동일 패턴
          google_event_id: string | null;        // Google Calendar 이벤트 ID (NULL=미동기화)
          google_event_updated_at: string | null; // 마지막 동기화 시각
          received_at: string;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          as_id: string;
          customer_id?: string | null;
          name: string;
          phone: string;
          proceed_type?: string | null;
          postcode?: string | null;
          address?: string | null;
          address_detail?: string | null;
          pickup_date?: string | null;
          delivery_method?: string | null;
          qty_mamoru?: number;
          qty_other?: number;
          memo?: string | null;
          service_cost?: number;
          shipping_fee?: number;
          total_amount?: number;
          status?: RepairStatus;
          invoice_number?: string | null;
          courier_name?: string | null;
          shipped_at?: string | null;
          shipped_source?: 'manual' | 'alps_pickup' | null;  // 109
          paid_at?: string | null;
          delivered_at?: string | null;
          confirmed_at?: string | null;
          packed_at?: string | null;
          admin_note?: string | null;
          gas_raw?: Record<string, unknown> | null;
          // 092: 직접방문(당일수리)
          visit_date?: string | null;
          visit_time?: string | null;
          visit_duration_min?: number | null;
          // 093: Google Calendar
          google_event_id?: string | null;
          google_event_updated_at?: string | null;
          received_at?: string;
        };
        Update: {
          as_id?: string;
          customer_id?: string | null;
          name?: string;
          phone?: string;
          proceed_type?: string | null;
          postcode?: string | null;
          address?: string | null;
          address_detail?: string | null;
          pickup_date?: string | null;
          delivery_method?: string | null;
          qty_mamoru?: number;
          qty_other?: number;
          memo?: string | null;
          service_cost?: number;
          shipping_fee?: number;
          total_amount?: number;
          status?: RepairStatus;
          invoice_number?: string | null;
          courier_name?: string | null;
          shipped_at?: string | null;
          shipped_source?: 'manual' | 'alps_pickup' | null;  // 109
          paid_at?: string | null;
          delivered_at?: string | null;
          confirmed_at?: string | null;
          packed_at?: string | null;
          admin_note?: string | null;
          gas_raw?: Record<string, unknown> | null;
          review_promised_at?: string | null;
          review_promised_type?: 'purchase' | 'repair' | 'consult' | null;
          review_promised_subtype?: string | null;
          review_request_sent_at?: string | null;
          review_submitted_at?: string | null;
          // 092: 직접방문(당일수리)
          visit_date?: string | null;
          visit_time?: string | null;
          visit_duration_min?: number | null;
          // 093: Google Calendar
          google_event_id?: string | null;
          google_event_updated_at?: string | null;
          received_at?: string;
        };
      };
      repair_inspections: {
        Row: {
          id: string;
          repair_id: string;
          scissor_number: number;
          scissor_type: string | null;
          blade_tip: string;
          blade_mid: string;
          blade_inner: string;
          comb: string;
          tension: string;
          parts: string;
          stopper: string;
          photo_url: string | null;
          photo_marks: Array<{ label: string; x?: number; y?: number; x2?: number; y2?: number; flag?: boolean }> | null;
          comment: string | null;  // 097: 가위별 진단 및 내역
          worker: string;
          created_at: string;
        };
        Insert: {
          repair_id: string;
          scissor_number: number;
          scissor_type?: string | null;
          blade_tip?: string;
          blade_mid?: string;
          blade_inner?: string;
          comb?: string;
          tension?: string;
          parts?: string;
          stopper?: string;
          photo_url?: string | null;
          photo_marks?: Array<{ label: string; x?: number; y?: number; x2?: number; y2?: number; flag?: boolean }> | null;
          comment?: string | null;
          worker?: string;
        };
        Update: {
          scissor_number?: number;
          scissor_type?: string | null;
          blade_tip?: string;
          blade_mid?: string;
          blade_inner?: string;
          comb?: string;
          tension?: string;
          parts?: string;
          stopper?: string;
          photo_url?: string | null;
          photo_marks?: Array<{ label: string; x?: number; y?: number; x2?: number; y2?: number; flag?: boolean }> | null;
          comment?: string | null;
          worker?: string;
        };
      };
      repair_history: {
        Row: {
          id: string;
          repair_id: string;
          from_status: RepairStatus | null;
          to_status: RepairStatus;
          changed_by: string | null;
          note: string | null;
          created_at: string;
        };
        Insert: {
          repair_id: string;
          from_status?: RepairStatus | null;
          to_status: RepairStatus;
          changed_by?: string | null;
          note?: string | null;
        };
        Update: {
          repair_id?: string;
          from_status?: RepairStatus | null;
          to_status?: RepairStatus;
          changed_by?: string | null;
          note?: string | null;
        };
      };
      consultation_history: {
        Row: {
          id: string;
          consultation_id: string;
          from_status: ConsultationStatus | null;
          to_status: ConsultationStatus;
          changed_by: string | null;
          note: string | null;
          created_at: string;
        };
        Insert: {
          consultation_id: string;
          from_status?: ConsultationStatus | null;
          to_status: ConsultationStatus;
          changed_by?: string | null;
          note?: string | null;
        };
        Update: {
          consultation_id?: string;
          from_status?: ConsultationStatus | null;
          to_status?: ConsultationStatus;
          changed_by?: string | null;
          note?: string | null;
        };
      };
      // R5: 오프라인 판매 테이블
      offline_sales: {
        Row: {
          id: string;
          sale_number: string;
          customer_id: string | null;
          customer_name: string;
          customer_phone: string | null;
          sale_date: string;
          total_amount: number;
          discount_amount: number;
          paid_amount: number;
          payment_method: PaymentMethod;
          payment_status: PaymentStatus;
          payment_detail: Record<string, number> | null; // 복합결제 분리 금액 {card,cash,transfer} (DB엔 있었으나 타입 누락 — 2026-08-06 보정)
          memo: string | null;
          supply_amount: number;
          vat_amount: number;
          is_vat_included: boolean;
          ecount_sync_status?: string | null;  // legacy
          ecount_slip_no?: string | null;       // legacy
          ecount_synced_at?: string | null;     // legacy
          cancelled_at: string | null;
          cancelled_reason: string | null;
          cancelled_by: string | null;
          returned_at: string | null;                // 반품 시점 (DB엔 있었으나 타입 누락 — 2026-08-02 보정)
          return_reason: string | null;              // 반품 사유
          sale_channel: SaleChannel;
          contract_id: string | null;
          review_requested_at: string | null;       // semantic alias of review_request_sent_at (legacy)
          review_promised_at: string | null;        // 067: 리뷰 약속 시점
          review_promised_type: 'purchase' | 'repair' | 'consult' | null; // 094: 리뷰 약속 유형
          review_promised_subtype: string | null; // 095: 약속 세부 유형 (direct_visit/pickup/store_visit/field_request/talk_consult)
          review_submitted_at: string | null;       // 067: 리뷰 작성 완료 시점
          source_consultation_id: string | null;    // 070: 출장/매장상담 → 판매 link (mirror 모드 트리거)
          // 038: 고객 유형 스냅샷 (B2C/B2B 판정 — dealer/academy = B2B)
          customer_type: string | null;
          // 048: 택배 발송 (송장 / 출고시각 / 택배사)
          invoice_number: string | null;
          shipped_at: string | null;
          courier_name: string | null;
          // 091: 배송완료/고객수령 시각 (cron ALPS 자동 OR "고객 수령 완료" 수동)
          delivered_at: string | null;
          // 109: 집하 자동감지 (출고 기록 주체 / 출고 알림톡 발송 시각)
          shipped_source: 'manual' | 'alps_pickup' | null;
          shipped_notified_at: string | null;
          // 111: 포장완료(준비완료) 시점 — 송장 무관, 내부 표시 전용 (2026-07-18)
          packed_at: string | null;
          created_by: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          sale_number: string;
          customer_id?: string | null;
          customer_name: string;
          customer_phone?: string | null;
          sale_date?: string;
          total_amount?: number;
          discount_amount?: number;
          paid_amount?: number;
          payment_method?: PaymentMethod;
          payment_status?: PaymentStatus;
          memo?: string | null;
          supply_amount?: number;
          vat_amount?: number;
          is_vat_included?: boolean;
          sale_channel?: SaleChannel;
          created_by?: string | null;
        };
        Update: {
          customer_id?: string | null;
          customer_name?: string;
          customer_phone?: string | null;
          sale_date?: string;
          total_amount?: number;
          discount_amount?: number;
          paid_amount?: number;
          payment_method?: PaymentMethod;
          payment_status?: PaymentStatus;
          memo?: string | null;
          supply_amount?: number;
          vat_amount?: number;
          is_vat_included?: boolean;
          cancelled_at?: string | null;
          cancelled_reason?: string | null;
          cancelled_by?: string | null;
          returned_at?: string | null;
          return_reason?: string | null;
          sale_channel?: SaleChannel;
          contract_id?: string | null;
          review_requested_at?: string | null;
          review_promised_at?: string | null;
          review_promised_type?: 'purchase' | 'repair' | 'consult' | null;
          review_promised_subtype?: string | null;
          review_submitted_at?: string | null;
          source_consultation_id?: string | null;
          customer_type?: string | null;
          // 048: 택배 발송
          invoice_number?: string | null;
          shipped_at?: string | null;
          courier_name?: string | null;
          // 091: 배송완료/수령완료
          delivered_at?: string | null;
          // 109: 집하 자동감지
          shipped_source?: 'manual' | 'alps_pickup' | null;
          shipped_notified_at?: string | null;
          // 111: 포장완료(준비완료) 시점
          packed_at?: string | null;
        };
      };
      offline_sale_items: {
        Row: {
          id: string;
          sale_id: string;
          product_id: string | null;
          product_name: string;
          sku: string | null;
          quantity: number;
          unit_price: number;
          total_price: number;
          category: string | null;
          supply_amount: number;
          vat_amount: number;
        };
        Insert: {
          sale_id: string;
          product_id?: string | null;
          product_name: string;
          sku?: string | null;
          category?: string | null;
          quantity?: number;
          unit_price?: number;
          total_price?: number;
          supply_amount?: number;
          vat_amount?: number;
        };
        Update: {
          product_id?: string | null;
          product_name?: string;
          sku?: string | null;
          quantity?: number;
          unit_price?: number;
          total_price?: number;
          supply_amount?: number;
          vat_amount?: number;
        };
      };
      // 빠른 송장(별도 송장) — 판매 무관 ALPS 송장 발급
      manual_invoices: {
        Row: {
          id: string;
          invoice_number: string;
          customer_id: string;
          customer_name: string;
          customer_phone: string;
          receiver_postcode: string;
          receiver_address_road: string;
          receiver_address_detail: string | null;
          goods_name: string;
          delivery_message: string | null;
          created_by: string | null;
          created_at: string;
          cancelled_at: string | null;
          cancelled_by: string | null;
          cancelled_reason: string | null;
          alps_cancel_failed: boolean;
        };
        Insert: {
          invoice_number: string;
          customer_id: string;
          customer_name: string;
          customer_phone: string;
          receiver_postcode: string;
          receiver_address_road: string;
          receiver_address_detail?: string | null;
          goods_name: string;
          delivery_message?: string | null;
          created_by?: string | null;
        };
        Update: {
          cancelled_at?: string | null;
          cancelled_by?: string | null;
          cancelled_reason?: string | null;
          alps_cancel_failed?: boolean;
        };
      };
      // R6: 전자 계약서 테이블
      contracts: {
        Row: {
          id: string;
          contract_number: string;
          customer_id: string | null;
          customer_name: string;
          customer_phone: string | null;
          customer_email: string | null;
          customer_address: string | null;
          total_amount: number;
          discount_amount: number;
          final_amount: number;
          payment_method: PaymentMethod;
          installment_months: number;
          signature_data: string | null;
          signed_at: string | null;
          pdf_url: string | null;
          image_url: string | null;
          notification_sent_at: string | null;
          status: ContractStatus;
          memo: string | null;
          ecount_sync_status?: string | null;  // legacy
          ecount_slip_no?: string | null;       // legacy
          offline_sale_id: string | null;
          delivery_method: string;
          unavailable_days: string | null;
          deposit_amount: number;
          balance_amount: number;
          seller_signature: string | null;
          customer_title: string | null;
          shop_name: string | null;
          shop_address: string | null;
          consultation_id: string | null;
          handwriting_name: string | null;
          handwriting_phone: string | null;
          handwriting_address: string | null;
          created_by: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          contract_number: string;
          customer_id?: string | null;
          customer_name: string;
          customer_phone?: string | null;
          customer_email?: string | null;
          customer_address?: string | null;
          total_amount?: number;
          discount_amount?: number;
          final_amount?: number;
          payment_method?: PaymentMethod;
          installment_months?: number;
          signature_data?: string | null;
          signed_at?: string | null;
          status?: ContractStatus;
          memo?: string | null;
          offline_sale_id?: string | null;
          delivery_method?: string;
          unavailable_days?: string | null;
          deposit_amount?: number;
          balance_amount?: number;
          seller_signature?: string | null;
          customer_title?: string | null;
          shop_name?: string | null;
          shop_address?: string | null;
          consultation_id?: string | null;
          handwriting_name?: string | null;
          handwriting_phone?: string | null;
          handwriting_address?: string | null;
          created_by?: string | null;
        };
        Update: {
          customer_id?: string | null;
          customer_name?: string;
          customer_phone?: string | null;
          customer_email?: string | null;
          customer_address?: string | null;
          total_amount?: number;
          discount_amount?: number;
          final_amount?: number;
          payment_method?: PaymentMethod;
          installment_months?: number;
          signature_data?: string | null;
          signed_at?: string | null;
          pdf_url?: string | null;
          image_url?: string | null;
          notification_sent_at?: string | null;
          status?: ContractStatus;
          memo?: string | null;
          offline_sale_id?: string | null;
          delivery_method?: string;
          unavailable_days?: string | null;
          deposit_amount?: number;
          balance_amount?: number;
          seller_signature?: string | null;
          customer_title?: string | null;
          shop_name?: string | null;
          shop_address?: string | null;
          consultation_id?: string | null;
          handwriting_name?: string | null;
          handwriting_phone?: string | null;
          handwriting_address?: string | null;
        };
      };
      contract_items: {
        Row: {
          id: string;
          contract_id: string;
          product_id: string | null;
          product_name: string;
          sku: string | null;
          quantity: number;
          unit_price: number;
          total_price: number;
          option_text: string | null;
        };
        Insert: {
          contract_id: string;
          product_id?: string | null;
          product_name: string;
          sku?: string | null;
          quantity?: number;
          unit_price?: number;
          total_price?: number;
          option_text?: string | null;
        };
        Update: {
          product_id?: string | null;
          product_name?: string;
          sku?: string | null;
          quantity?: number;
          unit_price?: number;
          total_price?: number;
          option_text?: string | null;
        };
      };
      // R7: 시리얼넘버 테이블
      product_serials: {
        Row: {
          id: string;
          product_id: string;
          serial_number: string;
          barcode: string | null;
          status: SerialStatus;
          sold_via: string | null;
          order_id: string | null;
          offline_sale_id: string | null;
          contract_id: string | null;
          sold_at: string | null;
          sold_to_name: string | null;
          sold_to_phone: string | null;
          lot_number: string | null;
          manufactured_at: string | null;
          memo: string | null;
          warehouse_zone: WarehouseZone;  // raw | ready | display
          created_by: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          product_id: string;
          serial_number: string;
          barcode?: string | null;
          status?: SerialStatus;
          sold_via?: string | null;
          order_id?: string | null;
          offline_sale_id?: string | null;
          contract_id?: string | null;
          sold_at?: string | null;
          sold_to_name?: string | null;
          sold_to_phone?: string | null;
          lot_number?: string | null;
          manufactured_at?: string | null;
          memo?: string | null;
          warehouse_zone?: string;
          created_by?: string | null;
        };
        Update: {
          serial_number?: string;
          barcode?: string | null;
          status?: SerialStatus;
          sold_via?: string | null;
          order_id?: string | null;
          offline_sale_id?: string | null;
          contract_id?: string | null;
          sold_at?: string | null;
          sold_to_name?: string | null;
          sold_to_phone?: string | null;
          lot_number?: string | null;
          manufactured_at?: string | null;
          memo?: string | null;
          warehouse_zone?: string;
        };
      };
      // Phase D: 매입(발주) 테이블
      purchase_orders: {
        Row: {
          id: string;
          po_number: string;
          supplier_id: string | null;
          supplier_name: string;
          order_date: string;
          expected_date: string | null;
          received_date: string | null;
          total_amount: number;
          deposit_amount: number;
          deposit_paid_at: string | null;
          balance_amount: number;
          balance_paid_at: string | null;
          is_vat_included: boolean;
          supply_amount: number;
          vat_amount: number;
          status: string; // draft | ordered | deposit_paid | received | balance_paid | cancelled
          memo: string | null;
          created_by: string | null;
          created_at: string;
          updated_at: string;
        };
        Insert: {
          po_number: string;
          supplier_id?: string | null;
          supplier_name: string;
          order_date?: string;
          expected_date?: string | null;
          total_amount?: number;
          deposit_amount?: number;
          balance_amount?: number;
          is_vat_included?: boolean;
          supply_amount?: number;
          vat_amount?: number;
          status?: string;
          memo?: string | null;
          created_by?: string | null;
        };
        Update: {
          supplier_name?: string;
          expected_date?: string | null;
          received_date?: string | null;
          total_amount?: number;
          deposit_amount?: number;
          deposit_paid_at?: string | null;
          balance_amount?: number;
          balance_paid_at?: string | null;
          supply_amount?: number;
          vat_amount?: number;
          status?: string;
          memo?: string | null;
        };
      };
      purchase_order_items: {
        Row: {
          id: string;
          po_id: string;
          product_id: string | null;
          product_name: string;
          sku: string | null;
          quantity: number;            // 주문 수량 (발주 기록 — 입고 시에도 덮어쓰지 않음)
          received_quantity: number | null;  // 081: 입고검수 실수령 수량. NULL = 입고 전
          unit_price: number;
          total_price: number;
        };
        Insert: {
          po_id: string;
          product_id?: string | null;
          product_name: string;
          sku?: string | null;
          quantity?: number;
          unit_price?: number;
          total_price?: number;
        };
        Update: {
          product_name?: string;
          sku?: string | null;
          quantity?: number;
          received_quantity?: number | null;
          unit_price?: number;
          total_price?: number;
        };
      };
    };
  };
}

/** 편의 타입 */
export type Profile = Database['public']['Tables']['profiles']['Row'];
export type Customer = Database['public']['Tables']['customers']['Row'];
export type Order = Database['public']['Tables']['orders']['Row'];
export type OrderItem = Database['public']['Tables']['order_items']['Row'];
export type Product = Database['public']['Tables']['products']['Row'];
/** 112: 창고 로케이션(정위치) */
export type WarehouseLocation = Database['public']['Tables']['warehouse_locations']['Row'];
/** 113: 렉 정보 */
export type WarehouseRack = Database['public']['Tables']['warehouse_racks']['Row'];
export type SyncLog = Database['public']['Tables']['sync_log']['Row'];

// Phase 2-1: 편의 타입
export type Dealer = Database['public']['Tables']['dealers']['Row'];
export type Consultation = Database['public']['Tables']['consultations']['Row'];
export type ConsultationHistory = Database['public']['Tables']['consultation_history']['Row'];

// Phase 7: 편의 타입
export type Repair = Database['public']['Tables']['repairs']['Row'];
export type RepairInspection = Database['public']['Tables']['repair_inspections']['Row'];
export type RepairHistory = Database['public']['Tables']['repair_history']['Row'];

// R5: 편의 타입
export type OfflineSale = Database['public']['Tables']['offline_sales']['Row'];
export type OfflineSaleItem = Database['public']['Tables']['offline_sale_items']['Row'];

// 빠른 송장
export type ManualInvoice = Database['public']['Tables']['manual_invoices']['Row'];

// R6: 편의 타입
export type Contract = Database['public']['Tables']['contracts']['Row'];
export type ContractItem = Database['public']['Tables']['contract_items']['Row'];

// R7: 편의 타입
export type ProductSerial = Database['public']['Tables']['product_serials']['Row'];

// Phase D: 편의 타입
export type PurchaseOrder = Database['public']['Tables']['purchase_orders']['Row'];
export type PurchaseOrderItem = Database['public']['Tables']['purchase_order_items']['Row'];
