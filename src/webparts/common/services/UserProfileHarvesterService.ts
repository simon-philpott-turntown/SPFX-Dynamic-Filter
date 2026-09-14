/**
 * @file UserProfileHarvesterService.ts
 * @description Shared dual-source harvester retrieving Graph v3 account attributes,
 * SharePoint PeopleManager properties, and profile photo blobs for authenticated users.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IUserProfileHarvestResult {
  details: Record<string, any>;
  photoUrl: string;
  firstName: string;
  displayName: string;
  isSiteAdmin: boolean;
}

export class UserProfileHarvesterService {
  private static _cache: IUserProfileHarvestResult | undefined;

  public static async harvest(context: WebPartContext): Promise<IUserProfileHarvestResult> {
    if (this._cache) {
      return this._cache;
    }

    const user = context?.pageContext?.user;
    const legacyPageContext = (window as any)._spPageContextInfo || (context?.pageContext as any)?.legacyPageContext;
    const isSiteAdmin = Boolean(legacyPageContext?.isSiteAdmin);

    const baseDetails: Record<string, any> = {
      DisplayName: user?.displayName || '',
      email: user?.email || '',
      loginName: user?.loginName || '',
      isSiteAdmin,
      isAnonymousGuestUser: Boolean(user?.isAnonymousGuestUser),
      isExternalGuestUser: Boolean(user?.isExternalGuestUser)
    };

    let photoUrl = '';
    const accountIdentifier = user?.email || user?.loginName || '';
    if (accountIdentifier && context?.pageContext?.web?.serverRelativeUrl) {
      const webUrl = context.pageContext.web.serverRelativeUrl === '/' ? '' : context.pageContext.web.serverRelativeUrl;
      photoUrl = `${webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(accountIdentifier)}`;
    }

    // 1. Microsoft Graph Harvester: standard + extended corporate attributes & manager
    try {
      if (context?.msGraphClientFactory) {
        const graphClient = await context.msGraphClientFactory.getClient('3');
        const graphUser = await graphClient
          .api('/me')
          .select('id,displayName,givenName,surname,userPrincipalName,mail,employeeId,department,companyName,jobTitle,officeLocation,city,state,country,postalCode,streetAddress,userType,onPremisesExtensionAttributes')
          .expand('manager($select=displayName,userPrincipalName)')
          .get();

        if (graphUser) {
          if (graphUser.givenName) baseDetails.givenName = graphUser.givenName;
          if (graphUser.givenName && !baseDetails.FirstName) baseDetails.FirstName = graphUser.givenName;
          if (graphUser.surname) baseDetails.surname = graphUser.surname;
          if (graphUser.jobTitle) baseDetails.jobTitle = graphUser.jobTitle;
          if (graphUser.department) baseDetails.department = graphUser.department;
          if (graphUser.companyName) baseDetails.companyName = graphUser.companyName;
          if (graphUser.employeeId) baseDetails.employeeId = graphUser.employeeId;
          if (graphUser.officeLocation) baseDetails.officeLocation = graphUser.officeLocation;
          if (graphUser.city) baseDetails.city = graphUser.city;
          if (graphUser.state) baseDetails.state = graphUser.state;
          if (graphUser.country) baseDetails.country = graphUser.country;
          if (graphUser.postalCode) baseDetails.postalCode = graphUser.postalCode;
          if (graphUser.streetAddress) baseDetails.streetAddress = graphUser.streetAddress;
          if (graphUser.manager?.displayName) baseDetails.manager = graphUser.manager.displayName;

          // Unpack onPremisesExtensionAttributes (extensionAttribute1..15)
          if (graphUser.onPremisesExtensionAttributes && typeof graphUser.onPremisesExtensionAttributes === 'object') {
            Object.keys(graphUser.onPremisesExtensionAttributes).forEach((extKey) => {
              const val = graphUser.onPremisesExtensionAttributes[extKey];
              if (val !== null && val !== undefined && val !== '') {
                baseDetails[extKey] = val;
              }
            });
          }
        }

        try {
          const timeoutPromise = new Promise<Blob>((_, reject) =>
            setTimeout(() => reject(new Error('Photo fetch timeout')), 1500)
          );
          const photoPromise: Promise<Blob> = graphClient
            .api('/me/photo/$value')
            .responseType('blob' as any)
            .get();

          const photoBlob = await Promise.race([photoPromise, timeoutPromise]);

          if (photoBlob && photoBlob.size > 0) {
            photoUrl = URL.createObjectURL(photoBlob);
          }
        } catch {
          // Keep SharePoint userphoto fallback
        }
      }
    } catch (graphErr) {
      console.warn('[UserProfileHarvesterService] Graph fetch error:', graphErr);
    }

    // 2. SharePoint User Profile Service Harvester (SP.UserProfiles.PeopleManager)
    try {
      if (context?.spHttpClient && context?.pageContext?.web?.absoluteUrl) {
        const endpoint = `${context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        );

        if (res && res.ok) {
          const data = await res.json();
          if (data) {
            if (data.DisplayName && !baseDetails.PreferredName) baseDetails.PreferredName = data.DisplayName;
            if (data.Department && !baseDetails.department) baseDetails.department = data.Department;
            if (data.Title && !baseDetails.jobTitle) baseDetails.jobTitle = data.Title;
            if (data.Office && !baseDetails.officeLocation) baseDetails.officeLocation = data.Office;

            const upsProps: Record<string, string> = {};
            if (data.WorkPhone) upsProps.workPhone = data.WorkPhone;

            if (Array.isArray(data.UserProfileProperties)) {
              data.UserProfileProperties.forEach((item: { Key?: string; Value?: string }) => {
                if (item && item.Key && item.Value && item.Value.trim() !== '') {
                  const technicalSkip = [
                    'SPS-FeedIdentifier', 'msOnline-ObjectId', 'PictureURL',
                    'SIPAddress', 'SPS-PrivacyActivity', 'SPS-PrivacyPeople', 'SPS-DistinguishedName'
                  ];
                  if (technicalSkip.indexOf(item.Key) === -1) {
                    if (item.Key === 'Department') upsProps.department = item.Value;
                    else if (item.Key === 'Title') upsProps.jobTitle = item.Value;
                    else if (item.Key === 'Office') upsProps.officeLocation = item.Value;
                    else if (item.Key === 'WorkPhone') upsProps.workPhone = item.Value;
                    else if (item.Key === 'FirstName' && !baseDetails.FirstName) upsProps.FirstName = item.Value;
                    else if (item.Key === 'PreferredName' && !baseDetails.PreferredName) upsProps.PreferredName = item.Value;
                    else if (item.Key === 'UserRegion' || item.Key.toLowerCase().includes('region')) upsProps.userRegion = item.Value;
                    else if (item.Key === 'EmployeeID' || item.Key === 'EmployeeId') upsProps.employeeId = item.Value;
                    else if (item.Key === 'Company' || item.Key === 'CompanyName') upsProps.companyName = item.Value;
                    else if (!baseDetails[item.Key]) {
                      upsProps[item.Key] = item.Value;
                    }
                  }
                }
              });
            }

            Object.assign(baseDetails, upsProps);
          }
        }
      }
    } catch (upsErr) {
      console.warn('[UserProfileHarvesterService] PeopleManager fetch error:', upsErr);
    }

    const firstName = baseDetails.FirstName || baseDetails.givenName || (user?.displayName ? user.displayName.split(' ')[0] : 'there');
    const displayName = baseDetails.PreferredName || baseDetails.DisplayName || user?.displayName || 'Colleague';

    const result: IUserProfileHarvestResult = {
      details: baseDetails,
      photoUrl,
      firstName,
      displayName,
      isSiteAdmin
    };

    this._cache = result;
    return result;
  }
}
