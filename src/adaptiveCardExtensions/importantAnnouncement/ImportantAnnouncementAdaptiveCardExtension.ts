import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseAdaptiveCardExtension } from "@microsoft/sp-adaptive-card-extension-base";
import { CardView } from "./cardView/CardView";
import { QuickView } from "./quickView/QuickView";
import { ImportantAnnouncementPropertyPane } from "./ImportantAnnouncementPropertyPane";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IImportantAnnouncementAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
}

export interface IImportantAnnouncementAdaptiveCardExtensionState {
  announcement: IAnnouncement | undefined;
}
export interface IAnnouncement {
  title: string;
  url: string;
}

const CARD_VIEW_REGISTRY_ID: string = "ImportantAnnouncement_CARD_VIEW";
export const QUICK_VIEW_REGISTRY_ID: string =
  "ImportantAnnouncement_QUICK_VIEW";

export default class ImportantAnnouncementAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IImportantAnnouncementAdaptiveCardExtensionProps,
  IImportantAnnouncementAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: ImportantAnnouncementPropertyPane | undefined;

  private _fetchAnnouncements(): Promise<void> {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Announcements')/items?$filter=Important eq 1&$select=Title,ID`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            accept: "application/json;odata.metadata=none",
          },
        }
      )
      .then((response) => response.json())
      .then((announcements) => {
        const announcement = announcements.value.pop();
        console.log(announcement);
        this.setState({
          announcement: {
            title: announcement.Title,
            url: `${this.context.pageContext.web.absoluteUrl}/lists/Announcements/DispForm.aspx?ID=${announcement.ID}`,
          },
        });
      })
      .catch((error) => console.error(error));
  }

  public onInit(): Promise<void> {
    this.state = {
      announcement: undefined,
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(
      QUICK_VIEW_REGISTRY_ID,
      () => new QuickView()
    );

    return this._fetchAnnouncements();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return "warning";
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'ImportantAnnouncement-property-pane'*/
      "./ImportantAnnouncementPropertyPane"
    ).then((component) => {
      this._deferredPropertyPane =
        new component.ImportantAnnouncementPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
