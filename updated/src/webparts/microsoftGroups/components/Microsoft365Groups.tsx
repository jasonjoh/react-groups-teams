import * as React from 'react';
import { MSGraphClient } from "@microsoft/sp-http";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './MicrosoftGroups.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { iconClass } from './MyTeams';
import { GroupDisplay } from './GroupDisplay';
import { Group, PlannerPlan } from '@microsoft/microsoft-graph-types';

export interface IGraphConsumerProps {
  context: WebPartContext;
}
export interface IUserItem {
  Topic: string;
  DeliveryDate: Date;
}

export interface IGraphConsumerState {
  AllGroups: GroupDisplay[];
  ShownGroups: GroupDisplay[];
  mode: string;
  title: string;
  isOpen: boolean;
  MoreDetails: any;
  Name: string;
  Description: string;
}

export default class MicrosoftGroups extends React.Component<IGraphConsumerProps, IGraphConsumerState> {
  private graphClient: MSGraphClient = null;
  public Tenant = this.props.context.pageContext.web.absoluteUrl.split('.')[0].split('//')[1];

  constructor(props) {
    super(props);
    this.GetPlanner = this.GetPlanner.bind(this);
    this.GetGroups = this.GetGroups.bind(this);
    this.state = {
      AllGroups: [],
      ShownGroups: [],
      mode: 'All',
      title: 'Groups In My Organization',
      isOpen: false,
      MoreDetails: [],
      Name: '',
      Description: ''
    };
  }

  private GetGroupsToShow(includeOnlyUser: boolean, visibility: string): GroupDisplay[] {
    if (includeOnlyUser) {
      if (visibility === 'All') {
        return this.state.AllGroups.filter(group => group.IsUserMember);
      } else {
        return this.state.AllGroups.filter(group => group.Visibility === visibility && group.IsUserMember);
      }
    } else {
      return this.state.AllGroups.filter(group => group.Visibility === visibility);
    }
  }

  public SwitchGroupList() {
    if (this.state.title === 'Groups In My Organization') {
      const myGroups = this.GetGroupsToShow(true, this.state.mode);
      this.setState({
        title: 'My Groups',
        ShownGroups: myGroups
      });
    }
    else {
      const allGroups = this.GetGroupsToShow(false, this.state.mode);
      this.setState({
        title: 'Groups In My Organization',
        ShownGroups: allGroups
      });
    }
  }

  public SwitchGroupList2(Switch) {
    const showMyGroups = this.state.title === 'My Groups';
    const groupsToShow = this.GetGroupsToShow(showMyGroups, Switch);

    this.setState({
      mode: Switch,
      ShownGroups: groupsToShow
    });
  }

  public OpenModal(GroupInfo) {
    var array = [];
    array.push(GroupInfo);
    this.setState({ isOpen: true, MoreDetails: array, Name: GroupInfo.Name, Description: GroupInfo.Description });
  }

  public async GetPlanner(groupId: string): Promise<string> {
    const plans = await this.graphClient
      .api(`/groups/${groupId}/planner/plans`)
      .get();

    if (plans.value.length > 0) {
      var PlanID;

      // Note: Groups can have more than one plan, this
      // just picks the last one for simplicity's sake
      plans.value.map((plan: PlannerPlan) => {
        PlanID = plan.id;
      });

      return `https://tasks.office.com/${this.Tenant}.com/EN-US/Home/Planner#/plantaskboard?groupId=${groupId}&planId=${PlanID}`;
    }
  }

  public async GetGroups() {
    const allGroupsArray: GroupDisplay[] = [];

    try {
      // Get all groups in the org
      const allGroups = await this.graphClient
        .api('groups')
        .filter(`groupTypes/any(a:a eq 'unified')`)
        .get();

      // Get the user's joined groups
      const joinedGroups = await this.graphClient
        //.api(`me/transitiveMemberOf/microsoft.graph.group?$filter=groupTypes/any(a:a eq 'unified')&$select=id`)
        .api('me/transitiveMemberOf/microsoft.graph.group')
        .filter(`groupTypes/any(a:a eq 'unified')`)
        .select('id')
        .get();

      await Promise.all(allGroups.value.map(async (group: Group) => {
        // If this group's ID matches an ID in the user's
        // groups, then the user is a member
        const isUserMember = joinedGroups.value
          .find((myGroup: Group) => { return myGroup.id === group.id; }) !== undefined;

        const planner = await this.GetPlanner(group.id);

        allGroupsArray.push({
          Name: group.displayName,
          Id: group.id,
          Description: group.description,
          Mail: group.mail,
          Visibility: group.visibility,
          IsUserMember: isUserMember,
          Planner: planner
        });
      }));

      this.setState({ AllGroups: allGroupsArray, ShownGroups: allGroupsArray });
    } catch (err) {
      console.log(JSON.stringify(err));
      throw err;
    }
  }

  public async componentDidMount() {
    // Get the Graph client once here
    try {
      this.graphClient = await this.props.context.msGraphClientFactory.getClient();
    } catch (err) {
      console.log(JSON.stringify(err));
    }

    this.GetGroups();
  }

  public render(): React.ReactElement<IGraphConsumerProps> {
    var Replaceregex = /\s+/g;
    return <div className={styles.test}>
      <div className={styles.tableCaptionStyle}>{this.state.title}
        <div>
          {this.state.mode === 'Public' ? <button className={styles.SelectedFilter} onClick={() => this.SwitchGroupList2('Public')}>Public</button> :
          <button className={styles.Filters} onClick={() => this.SwitchGroupList2('Public')}>Public</button>}

          {this.state.mode === 'All' ? <button className={styles.SelectedFilter} onClick={() => this.SwitchGroupList2('All')}>All</button> :
          <button className={styles.Filters} onClick={() => this.SwitchGroupList2('All')}>All</button>}

          {this.state.mode === 'Private' ? <button className={styles.SelectedFilter} onClick={() => this.SwitchGroupList2('Private')}>Private</button> :
          <button className={styles.Filters} onClick={() => this.SwitchGroupList2('Private')}>Private</button>}

        </div>
        <button className={styles.SwitchGroups} onClick={() => this.SwitchGroupList()}> View {this.state.title === 'My Groups' ?
        'Groups in my Organization' : 'My Groups'}</button>
      </div>
      <div className={styles.tableStyle}>
        <div className={styles.headerStyle}>
          <div className={styles.Center}>Group</div>
          <div className={styles.Center}>Mail</div>
          <div className={styles.Center}>Site</div>
          <div className={styles.Center}>Calendar</div>
          <div className={styles.Center}>Planner</div>
          <div className={styles.Center} style={{ borderRight: 'none' }}>Visibility</div>

        </div>
        {this.state.ShownGroups.map(group => {
          var GroupEmailSplit = group.Mail.split("@");
          group.Mail = GroupEmailSplit[0];
          return <div className={styles.rowStyle}>
            <div className={styles.ToolTipName}>{group.Name}<span className={styles.ToolTip}>{group.Description}</span></div>
            <a className={styles.Center} href={`https://outlook.office365.com/mail/group/${this.Tenant}.com/${group.Mail.toLowerCase()}/email`}>
              <Icon className={iconClass} style={{ color: '#087CD7' }} iconName="OutlookLogo"></Icon>
            </a>
            <a className={styles.Center} href={`https://${this.Tenant}.sharepoint.com/sites/${group.Mail}`}>
              <Icon className={iconClass} style={{ color: '#068B90' }} iconName="SharePointLogo"></Icon>
            </a>
            <a className={styles.Center} href={`https://outlook.office365.com/calendar/group/${this.Tenant}.com/${group.Name.replace(Replaceregex, '')}/view/week`}>
              <Icon className={iconClass} style={{ color: '#119AE2' }} iconName="Calendar"></Icon>
            </a>

            <div className={styles.Center}>
              {group.Planner === undefined ? <div></div> : <a href={group.Planner}>
                <Icon className={iconClass} style={{ color: '#077D3F' }} iconName="ViewListTree"></Icon></a>}
            </div>
            <div className={styles.Center} style={{ borderRight: 'none' }}>{group.Visibility}</div>
          </div>;
        })}
      </div>
    </div>;
  }
}
