import * as React from "react";

export const sendValueAsPercentCallout = (): JSX.Element => {
  return (
      <span>
        When true, the value will be divided by 100 when sent.
      </span>
  );
};

export const ratingIconCallout = (): JSX.Element => {
  return (
      <span>
        The <a href="https://developer.microsoft.com/fabric#/styles/icons">Office UI Fabric Icon</a> to use for the rating. Defaults to <i>FavoriteStar</i>.
      </span>
  );
};
