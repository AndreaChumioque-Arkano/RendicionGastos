import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";


import { List } from '@fluentui/react-teams';

class Tab extends React.Component {
  constructor(props){
    super(props)
    this.state = {
      context: {}
    }
  }

  componentDidMount(){
    // Get the user context from Teams and set it in the state
    microsoftTeams.getContext((context, error) => {
      this.setState({
        context: context
      });
    });
    // Next steps: Error handling using the error object
  }

  render() {
    const listProps = {
      find: true,
      filters: ["c2", "c3"],
      emptySelectionActionGroups: {
        g1: {
          a1: {
            title: "Nuevo",
            // icon: "Add",
          },
        },
      },
      columns: {
        c1: {
          title: "Member name",
          // sortable: "alphabetic",
        },
        c2: {
          title: "Location",
          hideable: true,
          minWidth: 100,
        },
        c3: {
          title: "Role",
          hideable: true,
          hidePriority: 1,
        },
      },
      rows: {
        r4: {
          c1: "Babak Shammas (no delete)",
          c2: "Seattle, WA",
          c3: "Senior analyst",
          actions: {
            share: { title: "Share", icon: "ShareGeneric" },
            manage: { title: "Edit", icon: "Edit" },
          },
        },
        r1: {
          c1: "Aadi Kapoor",
          c2: "Seattle, WA",
          c3: "Security associate",
          actions: {
            share: { title: "Share", icon: "ShareGeneric" },
            manage: { title: "Edit", icon: "Edit" },
            delete: { title: "Delete", icon: "TrashCan", multi: true },
          },
        },
        r2: {
          c1: "Aaron Buxton",
          c2: "Seattle, WA",
          c3:
            "Security engineer: Lorem ipsum dolor sit amet, consectetur adipiscing elit. Cras in ultricies mi. Sed aliquet odio et magna maximus, et aliquam ipsum faucibus. Sed pulvinar vel nibh eget scelerisque. Vestibulum ornare id felis ut feugiat. Ut vulputate ante non odio condimentum, eget dignissim erat tincidunt. Etiam sodales lobortis viverra. Sed gravida nisi at nisi ornare, non maximus nisi elementum.",
          actions: {
            share: { title: "Share", icon: "ShareGeneric" },
            manage: { title: "Edit", icon: "Edit" },
            delete: { title: "Delete", icon: "TrashCan", multi: true },
          },
        },
        r3: {
          c1: "Alvin Tao (no actions)",
          c2: "Seattle, WA",
          c3: "Marketing analyst",
        },
        r5: {
          c1: "Beth Davies",
          c2: "Seattle, WA",
          c3: "Senior engineer",
          actions: {
            share: { title: "Share", icon: "ShareGeneric" },
            manage: { title: "Edit", icon: "Edit" },
            delete: { title: "Delete", icon: "TrashCan", multi: true },
          },
        },
      },
    };

      let userName = Object.keys(this.state.context).length > 0 ? this.state.context['upn'] : "";

      return (
      <div>
        <h3>Hello World!</h3>
        <h1>Congratulations {userName}!</h1> <h3>This is the tab you made :-)</h3>
        <List
          {...listProps}
        />
      </div>
      );
  }
}
export default Tab;