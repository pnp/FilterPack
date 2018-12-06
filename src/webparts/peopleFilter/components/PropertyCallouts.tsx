import * as React from "react";

export const defaultValueCallout = (): JSX.Element => {
  return (
      <span>
        You can provide an email address (to provide multiple, just separate with a ;) that should be used for the starting value.
      </span>
  );
};

export const showHiddenInUICallout = (): JSX.Element => {
  return (
      <span>
        When true, users/groups that are hidden from the user interface will be included in results. This isn't usually needed.
      </span>
  );
};

export const groupNameCallout = (): JSX.Element => {
  return (
      <span>
        When specified, only members of the specified group are returned. Leave this blank to get all users.
      </span>
  );
};

export const useCurrentSiteCallout = (): JSX.Element => {
  return (
      <span>
        When true, users are pulled from this site. If you set it to false, you can specify a different site to pull users from (rarely needed).
      </span>
  );
};

export const webAbsoluteUrlCallout = (): JSX.Element => {
  return (
      <span>
        The URL of the site to pull users from. Include the full URL like: <br />
        <span style={{fontWeight:"bold",fontSize:"10px"}}>https://yourtenant.sharepoint.com/sites/yoursite</span>
      </span>
  );
};
