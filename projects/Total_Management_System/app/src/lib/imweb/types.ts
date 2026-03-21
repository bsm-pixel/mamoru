/** 아임웹 API 타입 — v2 + 새 OpenAPI 응답 구조 반영 */

/* ============================================================
 * 새 OpenAPI 상품 타입 (openapi.imweb.me)
 * ============================================================ */

export interface ImwebProduct {
  prodNo: number;
  name: string;
  prodStatus: string;          // 'sale' | 'soldout' | 'nosale' 등
  price: number;
  priceOrg: number;            // 원래 가격 (할인 전)
  stock: number;
  customSkuCode: string | null;
  imageUrl: string | null;
  categories: Array<{ categoryNo: number; categoryName: string }>;
  createTime: string;
  updateTime: string;
}

export interface ImwebProductListResponse {
  statusCode: number;
  data: {
    list: ImwebProduct[];
    totalCount: number;
    currentPage: number;
    lastPage: number;
  };
}

/* ============================================================
 * v2 API 주문 타입 (api.imweb.me)
 * ============================================================ */

export interface ImwebAuthResponse {
  access_token: string;
  expires_in: number;
}

export interface ImwebOrdersParams {
  order_date_from?: number;  // Unix timestamp
  order_date_to?: number;
  status?: string;
  page?: number;
  limit?: number;
}

export interface ImwebOrderItem {
  no: string;
  prod_no: string;
  prod_name: string;
  options: string;
  qty: number;
  price: number;
  total: number;
}

export interface ImwebOrder {
  order_no: string;
  order_code: string;
  order_time: number;
  order_type: string;
  complete_time: number;
  orderer: {
    member_code: string;
    name: string;
    email: string;
    call: string;           // phone → call
  };
  delivery: {
    country: string;
    address: {              // 중첩 address 객체
      name: string;
      phone: string;
      phone2: string;
      postcode: string;     // zipcode → postcode
      address: string;      // addr
      address_detail: string;
    };
    memo: string;
  };
  payment: {                // price → payment
    pay_type: string;
    pg_type: string;
    deliv_type: string;
    price_currency: string;
    total_price: number;
    deliv_price: number;
    coupon: number;
    payment_amount: number;
    payment_time: number;
  };
  items?: ImwebOrderItem[];
  parcel_code?: string;
  invoice_no?: string;
}

/** prod-orders 응답의 품목 */
export interface ImwebProdOrderItem {
  no: number;
  prod_no: number;
  prod_name: string;
  prod_custom_code: string | null;
  prod_sku_no: string;
  is_promotion: string;
  payment: {
    count: number;
    price: number;
    deliv_price: number;
    price_sale: number;
    coupon: number;
  };
}

/** prod-orders 응답 */
export interface ImwebProdOrder {
  order_no: string;        // "202602179557539-001" 형태
  status: string;          // PAY_COMPLETE, DELIVERY 등
  pay_time: number;
  items: ImwebProdOrderItem[];
}

export interface ImwebApiResponse<T> {
  code: number;
  msg: string;
  data: T;
}

export interface ImwebOrderListData {
  list: ImwebOrder[];
  pagenation: {
    data_count: number;
    current_page: number;
    total_page: number;
    pagesize: number;
  };
}
