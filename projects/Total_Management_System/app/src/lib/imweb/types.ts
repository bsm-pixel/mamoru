/** 아임웹 API v2 타입 */

export interface ImwebAuthResponse {
  access_token: string;
  expires_in: number;
}

export interface ImwebOrdersParams {
  order_date_from?: string;  // YYYY-MM-DD
  order_date_to?: string;
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
  status: string;
  pay_type: string;
  pay_time: number;
  order_time: number;
  orderer: {
    name: string;
    email: string;
    phone: string;
  };
  delivery: {
    name: string;
    phone: string;
    zipcode: string;
    addr: string;
    addr_detail: string;
    memo: string;
  };
  price: {
    total: number;
    deliv: number;
    discount: number;
    pay_price: number;
  };
  items: ImwebOrderItem[];
  parcel_code?: string;
  invoice_no?: string;
}

export interface ImwebApiResponse<T> {
  code: number;
  msg: string;
  data: T;
}

export interface ImwebOrderListData {
  list: ImwebOrder[];
  total: number;
  page: number;
  limit: number;
}
