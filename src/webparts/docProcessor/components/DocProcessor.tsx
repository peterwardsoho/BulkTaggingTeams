import * as React from "react";
import * as moment from "moment";
import styles from "./DocProcessor.module.scss";
import { IDocProcessorProps } from "./IDocProcessorProps";
import { IDocProcessorState } from "./IDocProcessorState";
import Dropzone from "react-dropzone";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp";
import {
  PrimaryButton,
  Dropdown,
  IDropdownOption,
} from "office-ui-fabric-react";
import {
  DateTimePicker,
  DateConvention,
} from "@pnp/spfx-controls-react/lib/dateTimePicker";
import TaxonomyWrapper from "./TaxonomyWrapper/TaxonomyWrapper";
import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { taxonomy, ITermStore, ITerm } from "@pnp/sp-taxonomy";

export default class DocProcessor extends React.Component<
  IDocProcessorProps,
  IDocProcessorState
> {
  constructor(props) {
    super(props);
    this.state = {
      allProjects: [],
      selectedProjectKey: "",
      selectedProjectText: "",
      uploads: [],
      submittalAction: [],
      gaPhase: [],
      agencies: [],
      documentStatus: [],
      byWhom: [],
      trades: [],
      fileNameTag: [],
      documentDate: null,
      actionDate: null,
      receivedDate: null,
      loading: false,
      loadingProjects: false,
      loadingScripts: true,
      errors: [],
    };
  }
  private getTermString(pickerTerms) {
    let termString = "";
    pickerTerms.forEach((term) => {
      termString += `-1;#${term["name"]}|${term["key"]};#`;
    });
    return termString.slice(0, -2);
  }
  // we need to set up the base URL of the site we are trying to deploy the webpart. 
  private loadProjects = async () => {
    sp.setup({
      sp: {
        baseUrl: "https://sohodragonlabs.sharepoint.com/sites/TestSite_BulkMetatagging_KTA",
      },
    });
    let projects = await sp.web.lists
      .getByTitle("Project List")
      .items.filter("Active_x0020_Status eq 'Active' and  MSTeam eq 1")
      .orderBy("Title")
      .get();
    this.setState({
      allProjects: projects,
    });
    console.log(projects);
  }
  public componentDidMount() {
    this.loadScripts();
    this.loadProjects();
  }
  public getSiteCollectionUrl(): string {
    let baseUrl = window.location.protocol + "//" + window.location.host;
    const pathname = window.location.pathname;
    const siteCollectionDetector = "/sites/";
    if (pathname.indexOf(siteCollectionDetector) >= 0) {
      baseUrl += pathname.substring(
        0,
        pathname.indexOf("/", siteCollectionDetector.length)
      );
    }
    return baseUrl;
  }
  private loadScripts() {
    const siteColUrl = this.getSiteCollectionUrl();
    try {
      SPComponentLoader.loadScript(siteColUrl + "/_layouts/15/init.js", {
        globalExportsName: "$_global_init",
      })
        .then(
          (): Promise<{}> => {
            return SPComponentLoader.loadScript(
              siteColUrl + "/_layouts/15/MicrosoftAjax.js",
              {
                globalExportsName: "Sys",
              }
            );
          }
        )
        .then(
          (): Promise<{}> => {
            return SPComponentLoader.loadScript(
              siteColUrl + "/_layouts/15/SP.Runtime.js",
              {
                globalExportsName: "SP",
              }
            );
          }
        )
        .then(
          (): Promise<{}> => {
            return SPComponentLoader.loadScript(
              siteColUrl + "/_layouts/15/SP.js",
              {
                globalExportsName: "SP",
              }
            );
          }
        )
        .then(
          (): Promise<{}> => {
            return SPComponentLoader.loadScript(
              siteColUrl + "/_layouts/15/SP.taxonomy.js",
              {
                globalExportsName: "SP",
              }
            );
          }
        )
        .then((): void => {
          this.setState({ loadingScripts: false });
        })
        .catch((reason: any) => {
          this.setState({
            loadingScripts: false,
            errors: [...this.state.errors, reason],
          });
        });
    } catch (error) {
      this.setState({
        loadingScripts: false,
        errors: [...this.state.errors, error],
      });
    }
  }
  private onDrop = (acceptedFiles) => {
    this.setState({
      uploads: acceptedFiles,
    });
  }
  private asyncForEach = async (array, callback) => {
    for (let index = 0; index < array.length; index++) {
      await callback(array[index], index, array);
    }
  }
  private setTaxFieldProperty = (
    context: SP.ClientContext,
    item: SP.ListItem,
    list: SP.List,
    fieldName: string,
    fieldValues: IPickerTerms,
    isSingle: boolean
  ) => {
    let field = list.get_fields().getByInternalNameOrTitle(fieldName);
    let taxField = context.castTo(
      field,
      SP.Taxonomy.TaxonomyField
    ) as SP.Taxonomy.TaxonomyField;

    if (isSingle) {
      if (fieldValues && fieldValues.length >= 1) {
        let termValue = new SP.Taxonomy.TaxonomyFieldValue();
        termValue.set_label(fieldValues[0].name);
        termValue.set_termGuid(new SP.Guid(fieldValues[0].key));
        termValue.set_wssId(-1);
        taxField.setFieldValueByValue(item, termValue);
      } else {
        taxField.validateSetValue(item, null);
      }
    } else {
      if (fieldValues && fieldValues.length >= 1) {
        let termString = this.getTermString(fieldValues);
        console.log(termString);
        let termValueCollection = new SP.Taxonomy.TaxonomyFieldValueCollection(
          context,
          termString,
          taxField
        );
        taxField.setFieldValueByValueCollection(item, termValueCollection);
      } else {
        taxField.validateSetValue(item, null);
      }
    }
  }
  private getItem = async (
    context: SP.ClientContext,
    list: SP.List,
    itemId: number
  ): Promise<SP.ListItem> => {
    return new Promise((resolve, reject) => {
      const query = `<View Scope='Recursive'><Query><Where><Eq><FieldRef Name=\"ID\" /><Value Type=\"Integer\">${itemId}</Value></Eq></Where></Query></View>`;
      let camlQuery = new SP.CamlQuery();
      camlQuery.set_viewXml(query);
      let allItems = list.getItems(camlQuery);
      context.load(allItems, "Include(Id)");
      context.executeQueryAsync(
        () => {
          console.log("success", allItems);
          resolve(allItems.get_item(0));
        },
        (sender, args) => {
          console.log("fail", args.get_message());
          resolve();
        }
      );
    });
  }
  private updateItemMetaData = async (
    siteUrl: string,
    listName: string,
    itemId: number
  ) => {
    console.log(siteUrl, listName, itemId);

    return new Promise(async (resolve, reject) => {
      let context = new SP.ClientContext(siteUrl);
      let list = context.get_web().get_lists().getByTitle(listName);
      console.log("ctx list", list);
      let item = await this.getItem(context, list, itemId);
      console.log("ctx item", item);

      this.setTaxFieldProperty(
        context,
        item,
        list,
        "ByWhom",
        this.state.byWhom,
        true
      );
      this.setTaxFieldProperty(
        context,
        item,
        list,
        "DocumentStatus",
        this.state.documentStatus,
        true
      );
      this.setTaxFieldProperty(
        context,
        item,
        list,
        "SubmittalAction",
        this.state.submittalAction,
        true
      );
      this.setTaxFieldProperty(
        context,
        item,
        list,
        "Phase",
        this.state.gaPhase,
        true
      );
      this.setTaxFieldProperty(
        context,
        item,
        list,
        "Agencies",
        this.state.agencies,
        true
      );
      this.setTaxFieldProperty(
        context,
        item,
        list,
        "Trades",
        this.state.trades,
        false
      );
      item.set_item("ActionDate", this.state.actionDate);

      item.update();
      context.load(item);

      context.executeQueryAsync(
        () => {
          console.log("success", item.get_id());
          resolve();
        },
        (sender, args) => {
          console.log("fail", args.get_message());
          resolve();
        }
      );
    });
  }
  private processFiles = async (projectSiteUrl) => {
    console.log(this.state);

    this.setState({
      loading: true,
    });

    //TODO: get info from project Taxonomy_aRm/OUPwaps1t+AbQGcHcQ==
    // const store: ITermStore = await taxonomy.termStores.getByName(
    //   "Taxonomy_aRm/OUPwaps1t+AbQGcHcQ=="
    // );
    // const projectTerm: ITerm = await store.getTermById(
    //   this.state.projects[0].key
    // );
    // const projectSiteUrl: string = await projectTerm.getDescription(1033);

    //switching up the context for the target site
    sp.setup({
      sp: {
        baseUrl: projectSiteUrl,
      },
    });

    await this.asyncForEach(this.state.uploads, async (file) => {
      let fileName = file.name;
      if (this.state.receivedDate) {
        let dateReceived = moment(this.state.receivedDate).format("YYYYMMDD");
        fileName = `${dateReceived}-${fileName}`;
      }
      if (this.state.fileNameTag && this.state.fileNameTag.length >= 1) {
        let fileNameTag = this.state.fileNameTag[0].name;
        fileName = `${fileNameTag}-${fileName}`;
      }
      console.log("file", file);
      console.log("fileName", fileName);

      //check if file exists, if so skip
      //check if file is checked out, if so skip
      console.log("pathTerm", this.state.gaPhase);
      const path = this.state.gaPhase[0].path;
      console.log("path", path);
      const pathTemp = path.split(";");
      console.log("pathTemp");
      const targetFolderName = pathTemp[0];
      const targetFolderUrl = targetFolderName.replace("-", "");
      console.log("Target Doc Library Name", targetFolderName, targetFolderUrl);

      const documentsFolderDisplayName = "Documents";
      const documentsFolderUrlName = "Shared%20Documents";

      let list = await sp.web.lists
        .getByTitle(documentsFolderDisplayName)
        .get();
      console.log("list", list);

      let fileUploaded = await sp.web
        .getFolderByServerRelativeUrl(
          `${documentsFolderUrlName}/${targetFolderName}`
        )
        .files.add(fileName, file, true); //using the url/folder name of the library
      console.log("file uploaded", fileUploaded);

      let listItemAllFields = await fileUploaded.file.listItemAllFields.get();
      console.log("list item", listItemAllFields);

      let update = await this.updateItemMetaData(
        projectSiteUrl,
        documentsFolderDisplayName,
        listItemAllFields.Id
      ); // using display list name of the library
      console.log(file.name + " properties updated successfully!");
    });
    console.log("processing files completed");
    this.reset();
  }
  private updateProjectSelection = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ) => {
    this.setState({
      loadingProjects: true
    }, ()=> {
      this.setState({
        selectedProjectKey: item.key as string,
        selectedProjectText: item.text
      }, () => {
        this.setState({
          loadingProjects: false
        });
      });
    });
  }
  private reset = () => {
    console.log("reseting form");
    this.setState({
      uploads: [],
      selectedProjectKey: "",
      selectedProjectText: "",
      submittalAction: [],
      gaPhase: [],
      agencies: [],
      documentStatus: [],
      byWhom: [],
      trades: [],
      fileNameTag: [],
      actionDate: null,
      receivedDate: null,
      loading: false,
    });
  }
  public render(): React.ReactElement<IDocProcessorProps> {
    const currentPhase =
      this.state.gaPhase.length > 0 ? this.state.gaPhase[0].path : "";
    const isCAPhase = currentPhase.indexOf("Contract Administration") != -1;
    let projects = [];
    this.state.allProjects.forEach((item) => {
      if (item["Project_x0020_Name"]) {
        projects.push({
          key: item["Project_x0020_Name"].Url,
          text: item.Title || item["Project_x0020_Name"].Description,
        });
      }
    });

    return (
      <div className={styles.docProcessor}>
        {!this.state.loadingScripts && (
          <div className={styles.row}>
            <div className={styles.left}>
              <Dropzone onDrop={this.onDrop} disabled={this.state.loading}>
                {({ getRootProps, getInputProps }) => (
                  <section>
                    <div {...getRootProps({ className: "dropzone" })}>
                      <input {...getInputProps()} />
                      <p>
                        Drag 'n' drop some files here, or click to select files
                      </p>
                    </div>
                  </section>
                )}
              </Dropzone>
              {this.state.uploads.map((upload) => (
                <div>{upload.name}</div>
              ))}
            </div>

            <div className={styles.right}>
              {!this.state.loading && (
                <Dropdown
                  className={styles.dropdown}
                  label="Project (required)"
                  required
                  onChange={this.updateProjectSelection}
                  options={projects}
                />
              )}
              {this.state.loading && (
                <Dropdown
                  className={styles.dropdown}
                  label="Project (required)"
                  required
                  disabled
                  onChange={this.updateProjectSelection}
                  options={projects}
                />
              )}

              {/* <TaxonomyWrapper
                allowMultipleSelections={false}
                termsetNameOrID="GA-Projects - Active"
                label="Project (required)"
                values={this.state.byWhom}
                context={this.props.context}
                onValueChanged={(pickerTerms) => {
                  this.setState({ projects: pickerTerms });
                }}
                loading={this.state.loading}
              /> */}

              <TaxonomyWrapper
                allowMultipleSelections={false}
                termsetNameOrID="GA-Phase"
                label="Phase (required)"
                context={this.props.context}
                values={this.state.gaPhase}
                onValueChanged={(pickerTerms) => {
                  console.log(pickerTerms);
                  this.setState({
                    gaPhase: pickerTerms,
                  });
                }}
                loading={this.state.loading}
              />

              <TaxonomyWrapper
                allowMultipleSelections={false}
                termsetNameOrID="By-Whom"
                label="By Whom"
                values={this.state.byWhom}
                context={this.props.context}
                onValueChanged={(pickerTerms) => {
                  this.setState({ byWhom: pickerTerms });
                }}
                loading={this.state.loading}
              />

              <TaxonomyWrapper
                allowMultipleSelections={false}
                termsetNameOrID="Document-Status"
                label="Document Status"
                context={this.props.context}
                values={this.state.documentStatus}
                onValueChanged={(pickerTerms) => {
                  this.setState({
                    documentStatus: pickerTerms,
                  });
                }}
                loading={this.state.loading}
              />

              {isCAPhase && (
                <TaxonomyWrapper
                  allowMultipleSelections={false}
                  termsetNameOrID="Submittal-Action"
                  label="Submittal Action (CA Phase Only)"
                  context={this.props.context}
                  values={this.state.submittalAction}
                  onValueChanged={(pickerTerms) => {
                    this.setState({
                      submittalAction: pickerTerms,
                    });
                  }}
                  loading={this.state.loading}
                />
              )}

              <TaxonomyWrapper
                allowMultipleSelections={true}
                termsetNameOrID="Trades"
                label="Trades"
                context={this.props.context}
                values={this.state.trades}
                onValueChanged={(pickerTerms) => {
                  this.setState({
                    trades: pickerTerms,
                  });
                }}
                loading={this.state.loading}
              />

              <TaxonomyWrapper
                allowMultipleSelections={false}
                termsetNameOrID="Agencies"
                label="Agencies"
                context={this.props.context}
                values={this.state.agencies}
                onValueChanged={(pickerTerms) => {
                  this.setState({
                    agencies: pickerTerms,
                  });
                }}
                loading={this.state.loading}
              />

              <DateTimePicker
                label="Action Date"
                dateConvention={DateConvention.Date}
                value={this.state.actionDate}
                showLabels={false}
                onChange={(date) => {
                  this.setState({ actionDate: date });
                }}
                disabled={this.state.loading}
              />

              <TaxonomyWrapper
                allowMultipleSelections={false}
                termsetNameOrID="FileName-Tag"
                label="File Name Tag"
                context={this.props.context}
                values={this.state.fileNameTag}
                onValueChanged={(pickerTerms) => {
                  this.setState({
                    fileNameTag: pickerTerms,
                  });
                }}
                loading={this.state.loading}
              />

              <DateTimePicker
                label="Received Date"
                dateConvention={DateConvention.Date}
                value={this.state.receivedDate}
                showLabels={false}
                onChange={(date) => {
                  this.setState({ receivedDate: date });
                }}
                disabled={this.state.loading}
              />

              <br />
              <PrimaryButton
                text="Process Files"
                onClick={() => this.processFiles(this.state.selectedProjectKey)}
                disabled={
                  this.state.loading ||
                  this.state.uploads.length <= 0 ||
                  this.state.selectedProjectKey.length <= 0 ||
                  this.state.gaPhase.length <= 0
                }
              />
            </div>
          </div>
        )}

        {this.state.errors}
      </div>
    );
  }
}
