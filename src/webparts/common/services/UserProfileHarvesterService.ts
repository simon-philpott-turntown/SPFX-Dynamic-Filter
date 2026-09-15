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
  private static readonly STORAGE_KEY = 'SPFX_HARVESTED_USER_PROFILE_v1';

  /**
   * Helper to construct the native SharePoint userphoto.aspx endpoint using absoluteUrl.
   */
  public static getUserPhotoUrl(context?: WebPartContext, size: 'S' | 'M' | 'L' = 'L'): string {
    if (!context) return '';
    const user = context.pageContext?.user;
    const accountIdentifier = user?.email || user?.loginName || '';
    if (!accountIdentifier) return '';

    const baseUrl = context.pageContext?.web?.absoluteUrl || window.location.origin;
    return `${baseUrl}/_layouts/15/userphoto.aspx?size=${size}&accountname=${encodeURIComponent(accountIdentifier)}`;
  }

  public static getCachedResult(context?: WebPartContext): IUserProfileHarvestResult | undefined {
    if (this._cache) {
      if (this._cache.photoUrl && this._cache.photoUrl.startsWith('blob:')) {
        this._cache.photoUrl = this.getUserPhotoUrl(context, 'L');
      }
      return this._cache;
    }

    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        const raw = window.localStorage.getItem(this.STORAGE_KEY);
        if (raw) {
          const parsed = JSON.parse(raw) as IUserProfileHarvestResult;
          if (parsed && parsed.details && typeof parsed.details === 'object') {
            // Scrub dead/stale blob: URLs from previous sessions
            if (parsed.photoUrl && parsed.photoUrl.startsWith('blob:')) {
              parsed.photoUrl = this.getUserPhotoUrl(context, 'L');
              try {
                window.localStorage.setItem(this.STORAGE_KEY, JSON.stringify(parsed));
              } catch {
                // Ignore storage update errors
              }
            } else if (!parsed.photoUrl && context) {
              parsed.photoUrl = this.getUserPhotoUrl(context, 'L');
            }
            this._cache = parsed;
            return parsed;
          }
        }
      }
    } catch {
      // Ignore localStorage read errors
    }
    return undefined;
  }

  public static async harvest(context: WebPartContext): Promise<IUserProfileHarvestResult> {
    const defaultUserPhoto = this.getUserPhotoUrl(context, 'L');

    const cached = this.getCachedResult(context);
    if (cached && cached.details && Object.keys(cached.details).length > 6) {
      if (!cached.photoUrl || cached.photoUrl.startsWith('blob:')) {
        cached.photoUrl = defaultUserPhoto;
      }
      this._cache = cached;
      // Continue background refresh non-blockingly if needed, but return cached immediately
    }

    const user = context?.pageContext?.user;
    const legacyPageContext = (window as any)._spPageContextInfo || (context?.pageContext as any)?.legacyPageContext;
    const isSiteAdmin = Boolean(legacyPageContext?.isSiteAdmin);

    const baseDetails: Record<string, any> = {
      ...(cached?.details || {}),
      DisplayName: user?.displayName || cached?.details?.DisplayName || '',
      email: user?.email || cached?.details?.email || '',
      loginName: user?.loginName || cached?.details?.loginName || '',
      isSiteAdmin,
      isAnonymousGuestUser: Boolean(user?.isAnonymousGuestUser),
      isExternalGuestUser: Boolean(user?.isExternalGuestUser)
    };

    let photoUrl = cached?.photoUrl || '';
    // Prevent using stale blob: URLs that may have been stored from previous sessions
    if (!photoUrl || photoUrl.startsWith('blob:')) {
      photoUrl = defaultUserPhoto;
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
            setTimeout(() => reject(new Error('Photo fetch timeout')), 2500)
          );
          const photoPromise: Promise<Blob> = graphClient
            .api('/me/photo/$value')
            .responseType('blob' as any)
            .get();

          const photoBlob = await Promise.race([photoPromise, timeoutPromise]);

          if (photoBlob && photoBlob.size > 0) {
            // Convert to base64 Data URL for persistent localStorage durability across page loads
            const base64Data = await this._blobToBase64(photoBlob);
            if (base64Data && base64Data.length > 50) {
              photoUrl = base64Data;
            }
          }
        } catch {
          // Graph photo not available or timed out: retain defaultUserPhoto
          if (!photoUrl || photoUrl.startsWith('blob:')) {
            photoUrl = defaultUserPhoto;
          }
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

            // Direct PictureUrl from PeopleManager only if valid http/https URL
            if (data.PictureUrl && typeof data.PictureUrl === 'string' && data.PictureUrl.trim().length > 0 && !data.PictureUrl.startsWith('blob:')) {
              if (!photoUrl || photoUrl === defaultUserPhoto) {
                photoUrl = data.PictureUrl.trim();
              }
            }

            const upsProps: Record<string, string> = {};
            if (data.WorkPhone) upsProps.workPhone = data.WorkPhone;

            if (Array.isArray(data.UserProfileProperties)) {
              data.UserProfileProperties.forEach((item: { Key?: string; Value?: string }) => {
                if (item && item.Key && item.Value && item.Value.trim() !== '') {
                  if (item.Key === 'PictureURL' || item.Key === 'PictureUrl') {
                    if ((!photoUrl || photoUrl === defaultUserPhoto) && !item.Value.startsWith('blob:')) {
                      photoUrl = item.Value.trim();
                    }
                    return;
                  }

                  const technicalSkip = [
                    'SPS-FeedIdentifier', 'msOnline-ObjectId',
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

    // Final safety check: if photoUrl is still empty or invalid, ensure defaultUserPhoto
    if (!photoUrl || photoUrl.startsWith('blob:')) {
      photoUrl = defaultUserPhoto;
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
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(this.STORAGE_KEY, JSON.stringify(result));
      }
    } catch {
      // Ignore localStorage write quota/privacy errors
    }

    return result;
  }

  /**
   * Serializes a binary image Blob into a Base64 Data URL for persistent storage in localStorage.
   */
  private static async _blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve) => {
      try {
        const reader = new FileReader();
        reader.onloadend = () => {
          resolve(typeof reader.result === 'string' ? reader.result : '');
        };
        reader.onerror = () => {
          resolve('');
        };
        reader.readAsDataURL(blob);
      } catch {
        resolve('');
      }
    });
  }
}
