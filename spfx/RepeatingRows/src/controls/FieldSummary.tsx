import * as React from 'react';
import { Text } from '@fluentui/react-components';

export function FieldSummary(props: { text: string }): React.ReactElement {
  return <Text>{props.text}</Text>;
}
