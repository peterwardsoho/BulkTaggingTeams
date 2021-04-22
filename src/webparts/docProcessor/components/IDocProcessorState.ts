import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

export interface IDocProcessorState {
  uploads: any[];
  allProjects: any[];
  selectedProjectKey: string;
  selectedProjectText: string;
  submittalAction: IPickerTerms;
  gaPhase: IPickerTerms;
  agencies: IPickerTerms;
  documentStatus: IPickerTerms;
  byWhom: IPickerTerms;
  trades: IPickerTerms;
  fileNameTag: IPickerTerms;
  documentDate: Date;
  actionDate: Date;
  receivedDate: Date;
  loading: boolean;
  loadingProjects: boolean;
  loadingScripts: boolean;
  errors: any;
}
