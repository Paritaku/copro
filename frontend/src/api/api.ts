import axiosInstance from "./axiosInstance";

const GET_TEMPLATE_MODELE_URL = "/modele";
const GET_FICHIERS_COPROPRIETE_URL = "/fichiers-copropriete";

export const getModele = async () => {
  try {
    const response = await axiosInstance.get(GET_TEMPLATE_MODELE_URL, {
      responseType: "blob",
    });
    return response.data;
  } catch (error) {}
};

export const getFichiersCopropriete = async (
  fichiersAGenerer: string[],
  file: File | null,
) => {
  if (file && fichiersAGenerer) {
    const formData = new FormData();
    formData.append("file", file);
    fichiersAGenerer.forEach((f) => formData.append("fichiersAGenerer", f));

    try {
      return await axiosInstance
        .post(
          GET_FICHIERS_COPROPRIETE_URL,
          formData,
          { responseType: "blob" },
        )
        .then((response) => response.data);
    } catch (error) {}
  }
};
