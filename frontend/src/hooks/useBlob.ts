export function useBlob(data: Blob, nom: string) {
  const url = URL.createObjectURL(data);
  const a = document.createElement("a");
  a.href = url;
  a.download = nom;
  document.body.appendChild(a);
  a.click();

  URL.revokeObjectURL(url);
  a.remove();
}
