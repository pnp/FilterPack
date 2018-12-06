import * as React from 'react';

export const defaultStateCallout = (): JSX.Element => {
    return (
        <span>
          This is the initial state of the toggle filter when the page loads.
        </span>
    );
};

export const sendValueAsStringCallout = (): JSX.Element => {
  return (
      <span>
        When true, the filter value will be sent as text (instead of true/false).
      </span>
  );
};
