import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import {
  SPHttpClient, SPHttpClientResponse
} from '@microsoft/sp-http';

export interface ISpScriptLinksApplicationCustomizerProperties {
}

export default class SpScriptLinksApplicationCustomizer
  extends BaseApplicationCustomizer<ISpScriptLinksApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    let requestUrl = webUrl.concat(`/_api/site/usercustomactions?$filter=Location eq 'ScriptLink' &$select=ScriptSrc,Location&$orderby=Sequence asc`);
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
    if (response.ok) {
      const res = await response.json();
      if (res && res.value && res.value.length && res.value.length > 0) {
        for (let uca of res.value) {
          if (uca.Location && uca.Location === 'ScriptLink' && uca.ScriptSrc && uca.ScriptSrc.length > 0) {
            const js = document.createElement('script');
            js.src = uca.ScriptSrc.toLowerCase()
              .replace('~sitecollection', siteUrl)
              .replace('~site', webUrl);
            const head = document.getElementsByTagName('head')[0];
            if (head && head.parentNode) {
              head.parentNode.insertBefore(js, head);
            }
          }
        }
      }
    }
    return Promise.resolve();
  }
}
