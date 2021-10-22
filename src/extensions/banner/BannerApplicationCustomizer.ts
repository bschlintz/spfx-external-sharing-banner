import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { sp, SPConfiguration, SearchQuery, SearchResults } from '@pnp/sp';

import * as strings from 'BannerApplicationCustomizerStrings';
import styles from './Banner.module.scss';

const LOG_SOURCE: string = 'ExternalSharingBanner.ApplicationCustomizer';
const CACHE_KEY: string = 'ExternalSharingBannerCache';

export interface IBannerApplicationCustomizerProperties {
  externalSharingText: string;
  textColor: string;
  backgroundColor: string;
  textFontSizePx: number;
  bannerHeightPx: number;
  cacheEnabled: boolean;
  cacheLifetimeMins: number;
  scope: 'site' | 'web';
  hiddenForExternalUsers: boolean;
  siteExclusionList: Array<string>;
}

const DEFAULT_PROPERTIES: IBannerApplicationCustomizerProperties = {
  externalSharingText: "This site contains <a href='{siteurl}/_layouts/15/siteanalytics.aspx'>content which is viewable by external users</a>",
  textColor: "#333",
  backgroundColor: "#ffffc6",
  textFontSizePx: 14,
  bannerHeightPx: 30,
  cacheEnabled: true,
  cacheLifetimeMins: 15,
  scope: 'site',
  hiddenForExternalUsers: true,
  siteExclusionList: []
};

/** A Custom Action which can be run during execution of a Client Side Application */
export default class BannerApplicationCustomizer
  extends BaseApplicationCustomizer<IBannerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent;
  private _extensionProperties: IBannerApplicationCustomizerProperties;
  private _bannerIsShowing: boolean = false;

  @override
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

      // Initialize PnP SP library
      sp.setup({ spfxContext: this.context } as SPConfiguration);

      // Merge passed properties with default properties, overriding any defaults
      this._extensionProperties = { ...DEFAULT_PROPERTIES, ...this.properties };

       // Don't show banner if the site is in the siteExclusionList config setting
       let url = this._extensionProperties.scope === 'web' ? this.context.pageContext.web.absoluteUrl : this.context.pageContext.site.absoluteUrl;
       if(this._extensionProperties.siteExclusionList.includes(url)){
         return;
       }
       
      // Don't show banner if the current user is an external user and the hiddenForExternalUsers config setting is enabled
      const isExternalUser = this.getIsExternalUser();
      if (isExternalUser && this._extensionProperties.hiddenForExternalUsers) {
        return;
      }

      // Event handler to re-render banner on each page navigation (detect if we change sites)
      this.context.application.navigatedEvent.add(this, this.onNavigated);
    });
  }

  @override
  public onDispose(): void {
    if (this._topPlaceholder) {
      this._topPlaceholder.dispose();
    }
  }

  /**
   * Event handler that fires on every page load
   */
  private async onNavigated(): Promise<void> {
    // Get external sharing status
    const isExternalSharingEnabled = await this.getExternalSharingStatus();
    // If external sharing is enabled, render banner, otherwise reset/clear banner if it was showing
    isExternalSharingEnabled ? this.renderBanner() : this.hideBanner();
  }

  /**
   * Get external sharing status for the current site collection from cache or REST API
   */
  private async getExternalSharingStatus(): Promise<boolean> {
    const { cacheLifetimeMins, scope } = this._extensionProperties;
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const isWebScope = scope === 'web';
    const cacheSubKey = `${siteId}_site${isWebScope ? `${webId}_web`: ``}`;

    // Check cache if it is enabled
    if (this._extensionProperties.cacheEnabled) {
      const cache = this.cacheGet(CACHE_KEY, cacheSubKey);
      if (null !== cache && undefined !== cache.isExternalSharingEnabled) {
        return cache.isExternalSharingEnabled;
      }
    }

    try {
      const isExternalSharingEnabled: boolean = await this.fetchExternalSharingStatus(siteId, isWebScope ? webId : undefined);

      // Check if cache is enabled, then cache result for the site with specified cache lifetime in minutes
      if (this. _extensionProperties.cacheEnabled) {
        this.cacheSet(CACHE_KEY, cacheSubKey, { isExternalSharingEnabled }, (1000 * 60 * cacheLifetimeMins));
      }

      return isExternalSharingEnabled;
    }
    catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to retrieve current site from search REST API. Error: ${error}`));

      return false;
    }
  }

  /**
   * Fetch external sharing status from Search REST API by checking if any content exists on the site that is viewable by external users
   * @param siteId Unique ID of the site
   * @param webId Optionally, get external sharing status for one web within a site
   */
  private async fetchExternalSharingStatus(siteId: string, webId?: string): Promise<boolean> {
    try {
      // Construct search query using PnP SP
      // Use Querytemplate so we don't clutter native search analytics
      const queryTemplateConditions = [
        `{searchterms}`,
        `SiteId:"${siteId}"`,
        `${webId ? `WebId:"${webId}"` : ``}`,
        `(ViewableByExternalUsers:1 OR ViewableByAnonymousUsers:1)`,
        `(IsDocument:1 OR ContentTypeId:0x010100* OR ContentTypeId:0x0120*)`,
        `-SecondaryFileExtension=one NOT IsOneNotePage=1`,
        `-ContentClass=STS_ListItem_GenericList`,
        `-ContentClass=STS_List_*`,
        `-ContentClass=STS_Site`,
        `-ContentClass=STS_Web`,
        `-ContentClass=STS_ListItem_544`,
        `-ContentClass=STS_ListItem_550`
      ];

      const searchQuery: SearchQuery = {
        Querytext: '*',
        QueryTemplate: queryTemplateConditions.join(' '),
        SelectProperties: ['Title', 'SiteID', 'ViewableByExternalUsers'],
        ClientType: 'Custom',
        RowLimit: 1,
        EnableQueryRules: false,
        SourceId: '8413cd39-2156-4e00-b54d-11efd9abdb89' // Local SharePoint Results
      };
      const results: SearchResults = await sp.search(searchQuery);

      //If we got any results back, then the site has content which is viewable by external users
      return results.TotalRows > 0;
    }
    catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to retrieve data from search REST API. Error: ${error}`));
      return false;
    }
  }

  /**
   * Check if the current user is an external user
   */
  private getIsExternalUser(): boolean {
    return this.context.pageContext.user.isExternalGuestUser || this.context.pageContext.user.isAnonymousGuestUser;
  }

  /**
   * Render the 'content viewable by external users' banner on the current page
   */
  private renderBanner(): void {
    const { externalSharingText, backgroundColor, textColor, textFontSizePx, bannerHeightPx } = this._extensionProperties;

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

      if (!this._topPlaceholder) {
        Log.error(LOG_SOURCE, new Error(`Unable to render Top placeholder`));
        return;
      }
    }

    //Only re-render banner if it isn't already showing
    if (!this._bannerIsShowing) {
      const textStyles = [
        `color: ${textColor};`,
        `font-size: ${textFontSizePx}px;`
      ];
      this._topPlaceholder.domElement.innerHTML = `
        <div style="background-color: ${backgroundColor};">
          <div class="${styles.BannerTextContainer}">
            <span style="${textStyles.join(' ')}">${this.parseTokens(externalSharingText)}</span>
          </div>
        </div>
      `;

      //Animate expand
      setTimeout(() => {
        let elem = document.querySelector(`.${styles.BannerTextContainer}`);
        elem.className += ` ${styles.BannerTextContainerExpand}`;
        elem.setAttribute('style', `height: ${bannerHeightPx}px !important;`);
        this._bannerIsShowing = true;
      }, 1);
    }
  }

  /**
   * Hide the 'content viewable by external users' banner if it is showing on the current page
   */
  private hideBanner(): void {
    if (this._topPlaceholder) {
      //Animate collapse
      let elem = document.querySelector(`.${styles.BannerTextContainer}`);
      elem.className = elem.className.replace(` ${styles.BannerTextContainerExpand}`, '');
      elem.removeAttribute('style');
    }
    this._bannerIsShowing = false;
  }

  /**
   * Helper function to get an item by key and subkey from session storage
   */
  private cacheGet = (cacheKey, cacheSubKey) => {
    try {
      const valueStr = sessionStorage.getItem(cacheKey);
      if (valueStr) {
        const val = JSON.parse(valueStr);
        const subVal = val[cacheSubKey];
        if (subVal) {
          return !(subVal.expiration && Date.now() > subVal.expiration) ? subVal.payload : null;
        }
      }
      return null;
    }
    catch (error) {
      Log.warn(LOG_SOURCE, `Unable to get cached data from sessionStorage with key ${cacheKey}. Error: ${error}`);
      return null;
    }
  }

  /**
   * Helper function to set an item by key and subkey into session storage
   */
  private cacheSet = (cacheKey, cacheSubKey, payload, expiresIn) => {
    try {
      const nowTicks = Date.now();
      const expiration = (expiresIn && nowTicks + expiresIn) || null;
      const valueStr = sessionStorage.getItem(cacheKey) || "{}";
      let cache = JSON.parse(valueStr);
      cache[cacheSubKey] = { payload, expiration };
      sessionStorage.setItem(cacheKey, JSON.stringify(cache));
      return this.cacheGet(cacheKey, cacheSubKey);
    }
    catch (error) {
      Log.warn(LOG_SOURCE, `Unable to set cached data into sessionStorage with key ${cacheKey}. Error: ${error}`);
      return null;
    }
  }

  private parseTokens = (textWithTokens: string): string => {
    const tokens = [
      { token: '{siteurl}', value: this.context.pageContext.site.absoluteUrl },
      { token: '{weburl}', value: this.context.pageContext.web.absoluteUrl },
    ];

    const outputText = tokens.reduce((text, tokenItem) => {
      return text.replace(tokenItem.token, tokenItem.value);
    }, textWithTokens);

    return outputText;
  }
}
