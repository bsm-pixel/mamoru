/** 이미지를 maxSize 이하로 리사이징 후 File로 반환 */
export function resizeImage(
  file: File,
  maxSize = 1200,
  quality = 0.8
): Promise<File> {
  return new Promise((resolve, reject) => {
    // 이미지가 아니면 원본 반환
    if (!file.type.startsWith('image/')) {
      resolve(file);
      return;
    }

    const img = new Image();
    const url = URL.createObjectURL(file);

    img.onload = () => {
      URL.revokeObjectURL(url);

      let { width, height } = img;

      // 이미 작으면 원본 반환
      if (width <= maxSize && height <= maxSize) {
        resolve(file);
        return;
      }

      // 비율 유지 리사이징
      if (width > height) {
        height = Math.round(height * (maxSize / width));
        width = maxSize;
      } else {
        width = Math.round(width * (maxSize / height));
        height = maxSize;
      }

      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;

      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(file); return; }

      ctx.drawImage(img, 0, 0, width, height);

      canvas.toBlob(
        (blob) => {
          if (!blob) { resolve(file); return; }
          const resized = new File([blob], file.name, {
            type: 'image/jpeg',
            lastModified: Date.now(),
          });
          resolve(resized);
        },
        'image/jpeg',
        quality
      );
    };

    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('이미지 로드 실패'));
    };

    img.src = url;
  });
}
