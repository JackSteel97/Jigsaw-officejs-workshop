export async function downloadImageBase64(url: string) {
  return new Promise<string>(async (resolve, reject) => {
    const response = await fetch(url);
    const blob = await response.blob();
    const reader = new FileReader();
    reader.readAsDataURL(blob);
    reader.onloadend = () => {
      resolve(reader.result.toString());
    };
  });
}