/* 간편진단 타입 정의 */

export interface DiagnosisOption {
  id: string;
  label: string;
  desc: string;
  icon: string; // SVG 인라인 문자열 or 이모지
  gifUrl?: string;
  lottieUrl?: string;
}

export interface DiagnosisQuestion {
  id: string;
  label: string;
  question: string;
  sub: string;
  multiple: boolean;
  hasGif: boolean;
  condition: ((answers: DiagnosisAnswers) => boolean) | null;
  options: DiagnosisOption[];
}

export type DiagnosisAnswers = Record<string, string | string[]>;
