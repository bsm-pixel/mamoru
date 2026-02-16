/** Supabase Database 타입 — 수동 정의 (Phase 1) */

export type OrderStatus =
  | 'pay_wait'
  | 'pay_done'
  | 'preparing'
  | 'shipping'
  | 'delivered'
  | 'confirmed'
  | 'cancelled'
  | 'refund_request'
  | 'refunded';

export type UserRole = 'owner' | 'staff' | 'director';

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
          total_orders: number;
          total_spent: number;
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['customers']['Row'], 'id' | 'created_at' | 'updated_at' | 'phone_normalized' | 'total_orders' | 'total_spent'>;
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
          image_url: string | null;
          tags: Record<string, unknown> | null;
          stock_quantity: number;
          barcode: string | null;
          is_active: boolean;
          created_at: string;
          updated_at: string;
        };
        Insert: Omit<Database['public']['Tables']['products']['Row'], 'id' | 'created_at' | 'updated_at' | 'stock_quantity' | 'is_active'>;
        Update: Partial<Database['public']['Tables']['products']['Insert']>;
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
    };
  };
}

/** 편의 타입 */
export type Profile = Database['public']['Tables']['profiles']['Row'];
export type Customer = Database['public']['Tables']['customers']['Row'];
export type Order = Database['public']['Tables']['orders']['Row'];
export type OrderItem = Database['public']['Tables']['order_items']['Row'];
export type Product = Database['public']['Tables']['products']['Row'];
export type SyncLog = Database['public']['Tables']['sync_log']['Row'];
