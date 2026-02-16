/** 조건부 className 결합 */
export function cn(
  ...inputs: (string | number | boolean | undefined | null)[]
): string {
  return inputs.filter((x) => typeof x === 'string' && x.length > 0).join(' ');
}
