import type { MouseEvent } from 'react';

/**
 * 모달 배경(backdrop) 클릭으로 닫기 — **드래그로 글자를 선택하다 바깥에서 놓아도 닫히지 않는다.**
 *
 * 문제(사장님 지적, 2026-09-16 · 반복 발생):
 *   배경 div 에 `onClick={onClose}` 만 달면, 모달 **안에서** 드래그를 시작해 **바깥에서** 손을 떼는 순간
 *   click 이벤트가 mousedown·mouseup 의 공통 조상(=배경)에서 발생해 **모달이 닫힌다.**
 *   글을 복사하거나 지우려고 드래그할 때마다 작성 중인 내용이 날아갔다.
 *
 * 왜 반복됐나: 2026-08 에 `create-delivery-modal` 한 곳만 ref 방식으로 고쳤고,
 *   그 뒤 만들어진 모달들은 다시 `onClick={onClose}` 로 손수 짜여졌다(공용 처리가 없었다).
 *   → 이 헬퍼 하나로 통일한다. **새 모달도 반드시 이걸 쓴다.**
 *
 * 해결: mousedown 도 배경에서 시작했을 때만 닫는다.
 *   상태를 DOM(dataset)에 두므로 ref·useState 가 필요 없고, 중간에 리렌더가 나도 값이 유지된다.
 *
 * 사용법 — 배경 div 의 `onClick={onClose}` 를 그대로 대체한다:
 *   <div className="fixed inset-0 ..." {...backdropClose(onClose)}>
 *     <div onClick={(e) => e.stopPropagation()}> ...본문... </div>
 *   </div>
 */
export function backdropClose(onClose: () => void) {
  return {
    onMouseDown: (e: MouseEvent<HTMLElement>) => {
      // 배경 자신에서 눌렀는지 (본문에서 시작한 드래그면 false)
      e.currentTarget.dataset.backdropDown = String(e.target === e.currentTarget);
    },
    onMouseUp: (e: MouseEvent<HTMLElement>) => {
      // 배경 자신에서 뗐는지 (배경에서 눌러 본문 안에서 뗀 경우도 닫지 않는다)
      e.currentTarget.dataset.backdropUp = String(e.target === e.currentTarget);
    },
    onClick: (e: MouseEvent<HTMLElement>) => {
      const ds = e.currentTarget.dataset;
      const ok = ds.backdropDown === 'true' && ds.backdropUp === 'true' && e.target === e.currentTarget;
      ds.backdropDown = 'false';
      ds.backdropUp = 'false';
      // 누른 곳·뗀 곳·클릭 대상이 모두 배경일 때만 닫는다
      if (ok) onClose();
    },
  };
}
