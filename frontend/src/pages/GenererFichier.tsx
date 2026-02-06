import { getFichiersCopropriete, getModele } from "@/api/api";
import CustomBeadCrumb from "@/components/components/CustomBeadCrumb";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Play, Upload } from "lucide-react";
import { useState } from "react";

export default function GenererFichier() {
  const dowloadTemplate = async () => {
    const blob = await getModele();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.style.display = "none";
    a.href = url;
    a.download = "TB_template.xlsx";

    document.body.appendChild(a);
    a.click();

    URL.revokeObjectURL(url);
    a.remove();
  };
  const [file, setFile] = useState<File | null>(null);
  const fichiersConfig = [
    {
      key: "Quot P CH2",
      label:
        "Tableau de repartition des quots-parts et dimilieme d'indivision (Quot P CH2)",
    },
    {
      key: "TR-N",
      label: "Tableau détaillé des superficies par niveau (TR-N)",
    },
    {
      key: "TR-C",
      label:
        "Tableau récapitulatif des superficies totales par consistance (TR-C)",
    },
    {
      key: "TA",
      label: "Tableau des contenances de la coproprieté (TA)",
    },
    {
      key: "Voix",
      label: "Le nombre de voix des coproprietaires (Voix)",
    },
    {
      key: "PV",
      label: "PV de coproprieté (PV)",
    },
    {
      key: "Reglement",
      label: "Règlement",
    },
  ];
  const [fichiersSelectionnes, setFichiersSelectionnes] = useState<string[]>(
    fichiersConfig.map((f) => f.key),
  );

  const updateSelection = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.checked) {
      setFichiersSelectionnes([...fichiersSelectionnes, e.target.name]);
    } else {
      setFichiersSelectionnes(
        fichiersSelectionnes.filter((f) => f !== e.target.name),
      );
    }
  };

  const genererFichiers = async () => {
    const data = await getFichiersCopropriete(fichiersSelectionnes, file);
    const url = URL.createObjectURL(data);
    const a = document.createElement("a");
    a.style.display = "none";
    a.href = url;
    a.download = "Fichier_de_coproprieté";
    document.body.appendChild(a);
    a.click()

    URL.revokeObjectURL(url);
    a.remove()
  }

  return (
    <CustomBeadCrumb pageTitle="Générer les fichiers de coproprieté">
      <div className="flex flex-col gap-6">
        <Card>
          <CardHeader>
            <CardTitle>Description et conditions d'utilisations</CardTitle>

            <CardDescription>
              Cet outil permet de générer automatiquement les fichiers de
              copropriété{" "}
              <span className="font-bold">
                {" "}
                à partir d’un fichier CSV TB retraçant les modifications
                successives d’un titre foncier
              </span>{" "}
              .
            </CardDescription>

            <CardContent className="px-4">
              <Card>
                <CardHeader className="">
                  <CardTitle className="text-red-400">
                    📌Conditions d'utilisations
                  </CardTitle>
                  <CardDescription className="text-red-400">
                    Le respect de ces règles est nécessaire pour garantir une
                    génération correcte.
                  </CardDescription>
                </CardHeader>
                <CardContent className="space-y-2 text-sm text-muted-foreground">
                  <ul className="list-disc pl-8 space-y-1">
                    <li>Le fichier doit être au format CSV</li>
                    <li>
                      Le template TB fourni doit être utilisé sans modification
                    </li>
                    <li>Les colonnes ne doivent pas être renommées</li>
                    <li>
                      Chaque ligne représente une modification successive du
                      titre foncier
                    </li>
                    <li>Les données doivent être complètes et cohérentes</li>
                  </ul>
                </CardContent>
              </Card>
            </CardContent>
          </CardHeader>
        </Card>

        {/* 📥 Template */}

        <Card>
          <CardHeader>
            <CardTitle>Fichier modèle</CardTitle>
            <CardDescription>
              Le fichier Excel permet une saisie simple et structurée des
              données.
            </CardDescription>
          </CardHeader>

          <CardContent className="space-y-3 text-sm text-muted-foreground">
            <ul className="list-disc pl-4 space-y-1">
              <li>Téléchargez le fichier Excel modèle</li>
              <li>
                Remplacez les champs par celui du titre foncier dont l'étude est
                en cours en suivant la{" "}
                <span className="font-bold">
                  même logique (niveau =&gt; liste lots =&gt; totale superficie)
                </span>
              </li>
              <li>
                Enregistrez ensuite le fichier{" "}
                <span className="font-bold">
                  au format CSV (UTF-8) avec comme délimiteur{" "}
                  <span className="text-red-400">';'</span>
                </span>
              </li>
            </ul>

            <Button variant="outline" onClick={dowloadTemplate}>
              Télécharger le modèle Excel
            </Button>
          </CardContent>
        </Card>

        {/* 📤 Upload */}
        <Card>
          <CardHeader>
            <CardTitle>Importer le fichier TB</CardTitle>
            <CardDescription>
              Sélectionnez le fichier CSV TB complété.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="flex h-32 flex-col items-center justify-center gap-2 rounded-md border border-dashed text-sm text-muted-foreground">
              <Upload className="h-5 w-5" />
              <Button
                size="sm"
                variant="ghost"
                className="hover:bg-transparent hover:shadow-none"
              >
                <Input
                  type="file"
                  className="hover:bg-primary-foreground"
                  accept=".csv"
                  onChange={(e) => {
                    setFile(e.target.files?.[0] || null);
                  }}
                />
              </Button>
            </div>
          </CardContent>
        </Card>

        {/* 📂 Sélection des fichiers */}
        <Card>
          <CardHeader>
            <CardTitle>Fichiers à générer</CardTitle>
            <CardDescription>
              Sélectionnez les documents de copropriété à produire.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-3 grid md:grid-cols-2">
            {fichiersConfig.map(({ key, label }) => (
              <div key={key} className="flex items-center space-x-2">
                <input
                  type="checkbox"
                  name={key}
                  checked={!!fichiersSelectionnes.find((c) => c === key)}
                  onChange={updateSelection}
                />
                <span className="text-sm">{label}</span>
              </div>
            ))}
          </CardContent>
        </Card>

        {/* ▶️ Action finale */}
        <Card>
          <CardContent className="flex items-center justify-between">
            <div className="text-sm text-muted-foreground flex flex-col gap-2">
              <span>
                Vérifiez votre fichier et vos sélections avant de lancer la
                génération.
              </span>
              {!file && (<span className="text-red-400">
                Aucun fichier sélectionné
              </span>)}
              {file && !file.name.endsWith(".csv") && (<span className="text-red-400">
                Le fichier doit être au format <span className="font-bold">CSV</span> 
              </span>)}
              {fichiersSelectionnes.length === 0 && (<span className="text-red-400">
                Sélectionnez <span className="font-bold">au moins un</span> fichier à générer
              </span>)}
            </div>
            <Button
              disabled={
                !file ||
                !file.name.endsWith(".csv") ||
                fichiersSelectionnes.length === 0
              }
              onClick={genererFichiers}
              size="lg"
            >
              <Play className="mr-2 h-4 w-4" />
              Générer les fichiers
            </Button>
          </CardContent>
        </Card>
      </div>
    </CustomBeadCrumb>


  );
}
