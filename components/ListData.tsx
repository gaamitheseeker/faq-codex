import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IListItem {
  Id: number;
  Title: string;
  _x8cea_x554f_: string; // Answer internal name
}

interface IListDataProps {
  context: WebPartContext;
}

export const ListData: React.FC<IListDataProps> = ({ context }) => {
  const [items, setItems] = useState<IListItem[]>([]);

  useEffect(() => {
    const getItems = async (): Promise<void> => {
      const url =
        `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('YOUR_LIST_NAME')/items` +
        `?$select=Id,Title,_x8cea_x554f_`;

      const response = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        throw new Error(`Unable to load list items: ${response.statusText}`);
      }

      setItems(await response.json());
    };

    getItems().catch(console.error);
  }, [context]);

  return (
    <div>
      <h2>List Data</h2>

      <table>
        <thead>
          <tr>
            <th>Title</th>
            <th>Answer</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item) => (
            <tr key={item.Id}>
              <td>{item.Title}</td>
              <td>{item._x8cea_x554f_}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
