/* =========================================================
   MAMORU 제품 데이터 (products.js)

   ★ 수정 검색 키워드:
   - [수정문구] : 제품명, 가격, 팁 메시지 등 텍스트 수정
   - [수정이미지] : 제품 이미지 URL 수정
   - [수정아이콘] : 가위 종류별 아이콘 수정
   ========================================================= */

const PRODUCTS = [
  // ================= 블런트 (BL) =================
  {
    id: 'BL-001',
    name: '마모루 소프트 블런트 5.5',  /* [수정문구] 제품명 */
    priceNum: 350000,
    priceText: '350,000원',  /* [수정문구] 가격 표시 */
    image: 'https://cdn.imweb.me/upload/sample-bl-001.jpg',  /* [수정이미지] 제품 이미지 URL */
    detailUrl: 'https://mamoru.kr/shop/view/BL001',  /* 상품 상세페이지 URL */
    tags: {
      type: ['BL'],
      stage: ['DE', 'IN'],
      feel: ['FEEL_SOFT'],
      style: ['St_BACK', 'St_NONE'],
      habit: ['HAB_WET', 'HAB_NONE'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'BL-002',
    name: '마모루 라이트 블런트 5.5',  /* [수정문구] */
    priceNum: 280000,
    priceText: '280,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-bl-002.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/BL002',
    tags: {
      type: ['BL'],
      stage: ['CE', 'IN'],
      feel: ['FEEL_SOFT', 'FEEL_NONE'],
      style: ['St_GO', 'St_NONE'],
      habit: ['HAB_WET', 'HAB_DRY', 'HAB_NONE'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'BL-003',
    name: '마모루 파워 블런트 6.0',  /* [수정문구] */
    priceNum: 380000,
    priceText: '380,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-bl-003.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/BL003',
    tags: {
      type: ['BL'],
      stage: ['DE'],
      feel: ['FEEL_POWER'],
      style: ['St_GO'],
      habit: ['HAB_DRY', 'HAB_NONE'],
      gender: ['M']
    }
  },
  {
    id: 'BL-004',
    name: '마모루 클래식 블런트 5.5',  /* [수정문구] */
    priceNum: 320000,
    priceText: '320,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-bl-004.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/BL004',
    tags: {
      type: ['BL'],
      stage: ['DE', 'IN'],
      feel: ['FEEL_POWER', 'FEEL_NONE'],
      style: ['St_GO', 'St_BACK'],
      habit: ['HAB_WET', 'HAB_DRY', 'HAB_NONE'],
      gender: ['FM', 'M']
    }
  },

  // ================= 틴닝 (TH) =================
  {
    id: 'TH-001',
    name: '마모루 프리미엄 틴닝 25%',  /* [수정문구] */
    priceNum: 180000,
    priceText: '180,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-th-001.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/TH001',
    tags: {
      type: ['TH'],
      stage: ['DE', 'IN', 'CE'],
      thRatio: ['TH_25'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'TH-002',
    name: '마모루 스탠다드 틴닝 25%',  /* [수정문구] */
    priceNum: 120000,
    priceText: '120,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-th-002.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/TH002',
    tags: {
      type: ['TH'],
      stage: ['CE', 'IN'],
      thRatio: ['TH_25'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'TH-003',
    name: '마모루 정밀 틴닝 15%',  /* [수정문구] */
    priceNum: 160000,
    priceText: '160,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-th-003.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/TH003',
    tags: {
      type: ['TH'],
      stage: ['DE'],
      thRatio: ['TH_15'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'TH-004',
    name: '마모루 쾌속 틴닝 35%',  /* [수정문구] */
    priceNum: 150000,
    priceText: '150,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-th-004.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/TH004',
    tags: {
      type: ['TH'],
      stage: ['DE', 'IN'],
      thRatio: ['TH_35'],
      gender: ['FM', 'M']
    }
  },

  // ================= 장가위 (LO) =================
  {
    id: 'LO-001',
    name: '마모루 장가위 7.0 (블런트 겸용)',  /* [수정문구] */
    priceNum: 420000,
    priceText: '420,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-lo-001.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/LO001',
    tags: {
      type: ['LO'],
      stage: ['DE'],
      loUse: ['LO_BL'],
      feel: ['FEEL_SOFT', 'FEEL_POWER'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'LO-002',
    name: '마모루 싱글링 전용 장가위 7.0',  /* [수정문구] */
    priceNum: 380000,
    priceText: '380,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-lo-002.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/LO002',
    tags: {
      type: ['LO'],
      stage: ['DE'],
      loUse: ['LO_SING'],
      gender: ['FM', 'M']
    }
  },

  // ================= 슬라이싱 (SL) =================
  {
    id: 'SL-001',
    name: '마모루 슬라이싱 가위 6.0',  /* [수정문구] */
    priceNum: 280000,
    priceText: '280,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-sl-001.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/SL001',
    tags: {
      type: ['SL'],
      stage: ['DE', 'IN'],
      slWhy: ['SL_NEW'],
      gender: ['FM', 'M']
    }
  },
  {
    id: 'SL-002',
    name: '마모루 프로 슬라이싱 6.0',  /* [수정문구] */
    priceNum: 350000,
    priceText: '350,000원',  /* [수정문구] */
    image: 'https://cdn.imweb.me/upload/sample-sl-002.jpg',  /* [수정이미지] */
    detailUrl: 'https://mamoru.kr/shop/view/SL002',
    tags: {
      type: ['SL'],
      stage: ['DE'],
      slWhy: ['SL_SAME'],
      slSameWhy: ['SL_SAME_WHY_UNCOM', 'SL_SAME_WHY_UNCOM1'],
      gender: ['FM', 'M']
    }
  }
];

/* =========================================================
   가위 종류별 정보 + 팁 메시지
   ========================================================= */
const TYPE_INFO = {
  BL: {
    icon: '✂️',  /* [수정아이콘] 블런트 아이콘 */
    name: '블런트',  /* [수정문구] */
    tip: null  /* [수정문구] 팁 메시지 (null이면 표시 안함) */
  },
  TH: {
    icon: '〰️',  /* [수정아이콘] 틴닝 아이콘 */
    name: '틴닝',  /* [수정문구] */
    tip: '💡 틴닝은 한 번 구매로 평생 사용 가능해요. 욕심을 내셔도 괜찮습니다!'  /* [수정문구] */
  },
  LO: {
    icon: '📏',  /* [수정아이콘] 장가위 아이콘 */
    name: '장가위',  /* [수정문구] */
    tip: null  /* [수정문구] */
  },
  SL: {
    icon: '🌊',  /* [수정아이콘] 슬라이싱 아이콘 */
    name: '슬라이싱',  /* [수정문구] */
    tip: null  /* [수정문구] */
  }
};

/* =========================================================
   점수 기반 추천 로직

   ★ 점수 체계:
   - 경력 매칭: +10점
   - 커트 느낌 매칭: +5점
   - 커트 스타일 매칭: +5점
   - 커트 습관 매칭: +3점
   - 성별 매칭: +2점
   - 틴닝 감모량 매칭: +10점
   - 장가위 용도 매칭: +10점
   - 슬라이싱 구매동기 매칭: +5점
   - 슬라이싱 불만족 이유 매칭: +5점
   ========================================================= */
function getRecommendedProducts(diagnosis) {
  const userStage = diagnosis.Q_STAGE;
  const userTypes = diagnosis.Q_TYPE || [];

  const result = {};

  userTypes.forEach(type => {
    const products = PRODUCTS
      .filter(p => p.tags.type.includes(type))
      .map(product => {
        let score = 0;

        // 경력 매칭: +10점
        if (product.tags.stage?.includes(userStage)) {
          score += 10;
        }

        // 커트 느낌 매칭: +5점 (BL, LO)
        if (product.tags.feel?.includes(diagnosis.Q_FEEL)) {
          score += 5;
        }

        // 커트 스타일 매칭: +5점 (BL)
        if (product.tags.style?.includes(diagnosis.Q_STYLE)) {
          score += 5;
        }

        // 커트 습관 매칭: +3점 (BL)
        if (product.tags.habit?.includes(diagnosis.Q_HABIT)) {
          score += 3;
        }

        // 성별 매칭: +2점
        if (product.tags.gender?.includes(diagnosis.Q_GENDER)) {
          score += 2;
        }

        // 틴닝 감모량 매칭: +10점
        if (product.tags.thRatio?.includes(diagnosis.Q_TH_RATIO)) {
          score += 10;
        }

        // 장가위 용도 매칭: +10점
        if (product.tags.loUse?.includes(diagnosis.Q_LO_USE)) {
          score += 10;
        }

        // 슬라이싱 구매동기 매칭: +5점
        if (product.tags.slWhy?.includes(diagnosis.Q_SL_WHY)) {
          score += 5;
        }

        // 슬라이싱 불만족 이유 매칭: +5점
        if (product.tags.slSameWhy?.includes(diagnosis.Q_SL_SAME_WHY)) {
          score += 5;
        }

        return { ...product, score };
      })
      .sort((a, b) => b.score - a.score);  // 점수 높은 순 정렬

    result[type] = products;
  });

  return result;
}
