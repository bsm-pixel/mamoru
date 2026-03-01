'use client';

import { useState, useCallback, useMemo } from 'react';
import { QUESTIONS, LABEL_MAP } from '@/lib/diagnosis/data';
import type { DiagnosisAnswers, DiagnosisQuestion } from '@/lib/diagnosis/types';

/* ─── 보이는 질문 계산 ─── */
function getVisibleQuestions(answers: DiagnosisAnswers): DiagnosisQuestion[] {
  return QUESTIONS.filter((q) => {
    if (q.condition === null) return true;
    return q.condition(answers);
  });
}

/* ─── SVG 아이콘 렌더러 ─── */
function Icon({ svg, className }: { svg: string; className?: string }) {
  return (
    <span
      className={className}
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  );
}

/* ─────────────────────────────────────
   메인 컴포넌트
   ───────────────────────────────────── */
export default function DiagnosisPage() {
  const [step, setStep] = useState(0);
  const [answers, setAnswers] = useState<DiagnosisAnswers>({});
  const [showResult, setShowResult] = useState(false);
  const [animKey, setAnimKey] = useState(0); // 전환 애니메이션 키

  const visible = useMemo(() => getVisibleQuestions(answers), [answers]);
  const question = visible[step] as DiagnosisQuestion | undefined;
  const totalSteps = visible.length;
  const progress = totalSteps > 0 ? ((step + 1) / totalSteps) * 100 : 0;

  /* 현재 질문의 답변 */
  const currentAnswer = question ? answers[question.id] : undefined;
  const hasAnswer = question?.multiple
    ? Array.isArray(currentAnswer) && currentAnswer.length > 0
    : !!currentAnswer;

  /* 옵션 선택 */
  const handleSelect = useCallback(
    (qId: string, optId: string, isMultiple: boolean) => {
      setAnswers((prev) => {
        const next = { ...prev };
        if (isMultiple) {
          const arr = Array.isArray(next[qId]) ? [...(next[qId] as string[])] : [];
          const idx = arr.indexOf(optId);
          if (idx > -1) arr.splice(idx, 1);
          else arr.push(optId);
          next[qId] = arr;
        } else {
          next[qId] = optId;
        }
        return next;
      });
    },
    [],
  );

  /* 다음 */
  const handleNext = useCallback(() => {
    const nextStep = step + 1;
    const nextVisible = getVisibleQuestions(answers);
    if (nextStep >= nextVisible.length) {
      setShowResult(true);
    } else {
      setStep(nextStep);
      setAnimKey((k) => k + 1);
    }
  }, [step, answers]);

  /* 이전 */
  const handleBack = useCallback(() => {
    if (step > 0) {
      setStep(step - 1);
      setAnimKey((k) => k + 1);
    }
  }, [step]);

  /* 다시 진단 */
  const handleReset = useCallback(() => {
    setAnswers({});
    setStep(0);
    setShowResult(false);
    setAnimKey((k) => k + 1);
  }, []);

  /* ─── 결과 화면 ─── */
  if (showResult) {
    return <ResultView answers={answers} onReset={handleReset} />;
  }

  if (!question) return null;

  const isSelected = (optId: string) =>
    question.multiple
      ? Array.isArray(currentAnswer) && currentAnswer.includes(optId)
      : currentAnswer === optId;

  return (
    <div className="min-h-dvh bg-[#F2F2EA] text-[#181725] select-none">
      <div className="mx-auto max-w-[800px] px-6 py-10 md:px-8 md:py-12">
        {/* 헤더 */}
        <div className="flex items-center justify-between mb-2">
          <span className="text-lg font-semibold">간편진단</span>
          <span className="text-sm font-semibold text-[#D4613E] bg-[#D4613E]/10 px-3.5 py-1.5 rounded-full">
            {step + 1}/{totalSteps}
          </span>
        </div>

        {/* 프로그레스 바 */}
        <div className="mb-10 md:mb-12">
          <div className="w-full h-1.5 bg-[#181725]/[0.06] rounded-full overflow-hidden">
            <div
              className="h-full bg-[#D4613E] rounded-full transition-[width] duration-400 ease-out"
              style={{ width: `${progress}%` }}
            />
          </div>
        </div>

        {/* 질문 영역 — 애니메이션 키로 재마운트 */}
        <div key={animKey} className="animate-fadeIn">
          {/* 라벨 pill */}
          <span className="inline-block px-3.5 py-1.5 bg-[#D4613E]/10 rounded-full text-[#D4613E] text-xs font-semibold mb-5">
            {question.label}
          </span>

          {/* 질문 텍스트 */}
          <h2 className="text-2xl md:text-[32px] font-bold leading-snug mb-2 whitespace-pre-line">
            {question.question}
          </h2>
          {question.sub && (
            <p className="text-sm md:text-base text-[#6B6980] leading-relaxed mb-8">
              {question.sub}
            </p>
          )}
          {!question.sub && <div className="mb-8" />}

          {/* 옵션 목록 */}
          {question.hasGif ? (
            /* GIF 타입 — 2열 그리드 유지 */
            <div className="grid grid-cols-2 gap-3 md:gap-4">
              {question.options.map((opt) => {
                const sel = isSelected(opt.id);
                return (
                  <button
                    key={opt.id}
                    type="button"
                    onClick={() => handleSelect(question.id, opt.id, question.multiple)}
                    className={`flex flex-col rounded-xl overflow-hidden border transition-all duration-200 text-left
                      ${sel
                        ? 'bg-[#D4613E]/[0.08] border-[#D4613E] scale-[1.01]'
                        : 'bg-[#FAFAF5] border-[#181725]/[0.08] hover:bg-[#D4613E]/[0.04] hover:border-[#D4613E]/20'
                      } active:scale-[0.98]`}
                  >
                    {/* GIF/placeholder 영역 */}
                    <div className="w-full h-20 md:h-40 bg-[#181725]/[0.02] flex items-center justify-center">
                      <div className="flex flex-col items-center text-[#B5B3C2] text-xs">
                        <Icon svg={opt.icon} className="[&>svg]:w-7 [&>svg]:h-7 mb-1" />
                        <span>준비중</span>
                      </div>
                    </div>
                    {/* 텍스트 + 체크 */}
                    <div className="flex items-center gap-3 px-3 py-2.5 md:px-4 md:py-4">
                      <div className="flex-1 min-w-0">
                        <div className="text-sm md:text-base font-semibold">{opt.label}</div>
                        {opt.desc && (
                          <div className="text-xs md:text-sm text-[#6B6980]">{opt.desc}</div>
                        )}
                      </div>
                      <CheckCircle checked={sel} />
                    </div>
                  </button>
                );
              })}
            </div>
          ) : (
            /* 기본형 — PC 2열 / 모바일 1열 */
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 md:gap-4">
              {question.options.map((opt) => {
                const sel = isSelected(opt.id);
                return (
                  <button
                    key={opt.id}
                    type="button"
                    onClick={() => handleSelect(question.id, opt.id, question.multiple)}
                    className={`flex items-center gap-4 p-4 md:p-5 rounded-xl border transition-all duration-200 text-left
                      ${sel
                        ? 'bg-[#D4613E]/[0.08] border-[#D4613E] scale-[1.01]'
                        : 'bg-[#FAFAF5] border-[#181725]/[0.08] hover:bg-[#D4613E]/[0.04] hover:border-[#D4613E]/20'
                      } active:scale-[0.98]`}
                  >
                    {/* 아이콘 */}
                    <div
                      className={`w-11 h-11 md:w-[52px] md:h-[52px] flex-shrink-0 flex items-center justify-center rounded-lg [&>svg]:w-full [&>svg]:h-full
                        ${sel ? 'bg-[#D4613E]/[0.12]' : 'bg-[#181725]/[0.03]'}`}
                    >
                      <Icon svg={opt.icon} className="w-full h-full [&>svg]:w-full [&>svg]:h-full" />
                    </div>
                    {/* 텍스트 */}
                    <div className="flex-1 min-w-0">
                      <div className="text-[15px] md:text-[17px] font-semibold">{opt.label}</div>
                      {opt.desc && (
                        <div className="text-[13px] md:text-sm text-[#6B6980]">{opt.desc}</div>
                      )}
                    </div>
                    {/* 체크 */}
                    <CheckCircle checked={sel} />
                  </button>
                );
              })}
            </div>
          )}

          {/* 복수선택 힌트 */}
          {question.multiple && (
            <div className="flex items-center justify-center gap-1.5 mt-5 py-2.5 bg-[#D4613E]/[0.06] rounded-lg text-sm text-[#D4613E]">
              <svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" className="w-4 h-4">
                <path d="M9 18h6" /><path d="M10 21h4" /><path d="M12 2a7 7 0 0 0-4 12.7V17h8v-2.3A7 7 0 0 0 12 2z" />
              </svg>
              여러 개 선택 가능
            </div>
          )}
        </div>

        {/* 네비게이션 버튼 */}
        <div className="flex items-center justify-center gap-3 mt-10 md:mt-12 md:justify-end">
          {step > 0 && (
            <button
              type="button"
              onClick={handleBack}
              className="w-12 h-12 md:w-[52px] md:h-[52px] flex-shrink-0 flex items-center justify-center rounded-lg
                bg-[#181725]/[0.04] border border-[#181725]/[0.08] text-[#6B6980]
                hover:bg-[#D4613E]/[0.08] hover:border-[#D4613E]/30 hover:text-[#D4613E] transition-all"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" className="w-5 h-5">
                <path d="M20 12H4" /><path d="M10 6l-6 6 6 6" />
              </svg>
            </button>
          )}
          <button
            type="button"
            onClick={handleNext}
            disabled={!hasAnswer}
            className="flex-1 md:flex-none md:w-[180px] h-14 md:h-[52px] rounded-full
              bg-[#D4613E] text-[#F2F2EA] font-semibold text-base
              disabled:opacity-35 disabled:cursor-not-allowed
              hover:enabled:bg-[#B85232] hover:enabled:-translate-y-0.5
              active:enabled:scale-[0.97] transition-all"
          >
            다음
          </button>
        </div>
      </div>
    </div>
  );
}

/* ─── 체크 서클 ─── */
function CheckCircle({ checked }: { checked: boolean }) {
  return (
    <div
      className={`w-[22px] h-[22px] md:w-[26px] md:h-[26px] rounded-full flex-shrink-0 flex items-center justify-center border-2 transition-all
        ${checked
          ? 'bg-[#D4613E] border-[#D4613E]'
          : 'border-[#181725]/15'
        }`}
    >
      {checked && (
        <svg viewBox="0 0 24 24" fill="none" stroke="#F2F2EA" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" className="w-3 h-3">
          <polyline points="20 6 9 17 4 12" />
        </svg>
      )}
    </div>
  );
}

/* ─── 결과 화면 ─── */
function ResultView({
  answers,
  onReset,
}: {
  answers: DiagnosisAnswers;
  onReset: () => void;
}) {
  /* 답변한 질문만 필터 (결과에 표시할 쌍) */
  const answeredQuestions = useMemo(() => {
    const vis = getVisibleQuestions(answers);
    return vis.filter((q) => {
      const a = answers[q.id];
      if (Array.isArray(a)) return a.length > 0;
      return !!a;
    });
  }, [answers]);

  /* 답변 → 라벨 문자열 변환 */
  const getAnswerLabel = (qId: string): string => {
    const a = answers[qId];
    if (!a) return '';
    if (Array.isArray(a)) {
      return a.map((v) => LABEL_MAP[v] || v).join(', ');
    }
    return LABEL_MAP[a] || a;
  };

  return (
    <div className="min-h-dvh bg-[#F2F2EA] text-[#181725] select-none animate-fadeIn">
      <div className="mx-auto max-w-[800px] px-6 py-10 md:px-8 md:py-14">
        {/* 완료 배지 */}
        <div className="text-center mb-8">
          <span className="inline-flex items-center gap-1.5 px-4 py-2 bg-[#829E86]/[0.12] rounded-full text-[#6E8E73] text-sm font-semibold">
            <svg viewBox="0 0 24 24" fill="#6E8E73" stroke="none" className="w-4 h-4">
              <path d="M12 2l2 6.5L20 12l-6 3.5L12 22l-2-6.5L4 12l6-3.5z" />
              <circle cx="20" cy="4" r="1" />
              <circle cx="4" cy="20" r="0.75" />
            </svg>
            진단 완료
          </span>
        </div>

        {/* 결과 카드 — 질문-답변 쌍 테이블 */}
        <div className="bg-[#FAFAF5] border border-[#181725]/[0.08] rounded-xl overflow-hidden">
          {answeredQuestions.map((q, i) => (
            <div
              key={q.id}
              className={`flex items-center justify-between px-5 py-4 md:px-7 md:py-5
                ${i < answeredQuestions.length - 1 ? 'border-b border-[#181725]/[0.06]' : ''}`}
            >
              <span className="text-sm text-[#6B6980] font-medium flex-shrink-0 mr-4">
                {q.label}
              </span>
              <span className="text-sm md:text-base font-semibold text-right">
                {getAnswerLabel(q.id)}
              </span>
            </div>
          ))}
        </div>

        {/* 다시 진단하기 */}
        <div className="mt-10 flex justify-center">
          <button
            type="button"
            onClick={onReset}
            className="w-full max-w-[480px] h-14 rounded-full
              bg-[#D4613E] text-[#F2F2EA] font-semibold text-base
              hover:bg-[#B85232] hover:-translate-y-0.5
              active:scale-[0.97] transition-all"
          >
            다시 진단하기
          </button>
        </div>
      </div>
    </div>
  );
}
