import * as React from 'react';
import { MSGraphClient } from "@microsoft/sp-http";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './MicrosoftGroups.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Group, Team, PlannerPlan } from '@microsoft/microsoft-graph-types';
import { TeamDisplay } from './TeamDisplay';

export const iconClass = mergeStyles({
  fontSize: 32,
  height: 32,
  width: 32
});

export interface IMyTeamsProps {
  context: WebPartContext;
  hidden: Boolean;
}

export interface IUserItem {
  Topic: string;
  DeliveryDate: Date;
}

export interface IMyTeamsState {
  MyTeams: TeamDisplay[];
  ShownTeams: TeamDisplay[];
  mode: string;
  title: string;
}

export default class MyTeams extends React.Component<IMyTeamsProps, IMyTeamsState> {
  public Tenant = this.props.context.pageContext.web.absoluteUrl.split('.')[0].split('//')[1];
  private graphClient: MSGraphClient = null;

  constructor(props) {
    super(props);
    this.state = {
      MyTeams: [],
      ShownTeams: [],
      mode: 'All',
      title: 'Teams in Microsoft Teams In My Organization',
    };
  }
  public SwitchGroupList(Switch) {
    var displayTeams;

    if (Switch === 'All') {
      displayTeams = this.state.MyTeams;
    }
    else {
      displayTeams = this.state.MyTeams.filter(item => item.Visibility === Switch);
    }
    this.setState({
      mode: Switch,
      ShownTeams: displayTeams
    });
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

  public async GetTeamsURL(teamId: string): Promise<string> {
    var team: Team = await this.graphClient
      .api(`/teams/${teamId}`)
      .select('webUrl')
      .get();

    return team.webUrl;
  }

  public async GetMail(groupId: string): Promise<string> {
    const group: Group = await this.graphClient
      .api(`groups/${groupId}`)
      .get();

    return group.mail;
  }

  public async GetMyTeams() {
    try {
      const myTeams = await this.graphClient
      .api(`me/joinedTeams`)
      .get();

      const myTeamsArray: TeamDisplay[] = [];
      await Promise.all(myTeams.value.map(async (team: Team) => {
        const mail = await this.GetMail(team.id);
        const planner = await this.GetPlanner(team.id);
        const teamUrl = await this.GetTeamsURL(team.id);

        const teamDisplay: TeamDisplay = {
          Name: team.displayName,
          Id: team.id,
          Description: team.description,
          Visibility: team.visibility,
          Mail: mail,
          Planner: planner,
          WebUrl: teamUrl
        };

        myTeamsArray.push(teamDisplay);
      }));

      this.setState({ MyTeams: myTeamsArray, ShownTeams: myTeamsArray });

    } catch (err) {
      console.log(JSON.stringify(err));
    }
  }

  public async componentDidMount() {
    // Get the Graph client once here
    try {
      this.graphClient = await this.props.context.msGraphClientFactory.getClient();
    } catch (err) {
      console.log(JSON.stringify(err));
    }

    this.GetMyTeams();
  }

  public render(): React.ReactElement<IMyTeamsProps> {
    var Replaceregex = /\s+/g;
    return this.props.hidden ? <div></div> : <div className={styles.test}>
      <div className={styles.tableCaptionStyle} style={{ borderRight: 'none' }}>My Teams Teams<div>

        {this.state.mode === 'Public' ? <button className={styles.SelectedFilter} onClick={() => this.SwitchGroupList('Public')}>Public</button> :
          <button className={styles.Filters} onClick={() => this.SwitchGroupList('Public')}>Public</button>}

        {this.state.mode === 'All' ? <button className={styles.SelectedFilter} onClick={() => this.SwitchGroupList('All')}>All</button> :
          <button className={styles.Filters} onClick={() => this.SwitchGroupList('All')}>All</button>}

        {this.state.mode === 'Private' ? <button className={styles.SelectedFilter} onClick={() => this.SwitchGroupList('Private')}>Private</button> :
        <button className={styles.Filters} onClick={() => this.SwitchGroupList('Private')}>Private</button>}</div>

      </div>
      <div className={styles.tableStyle}>
        <div className={styles.headerStyle}>
          <div className={styles.Center}>Team</div>
          <div className={styles.Center}>Mail</div>
          <div className={styles.Center}>Site</div>
          <div className={styles.Center}>Calendar</div>
          <div className={styles.Center}>Planner</div>
          <div className={styles.Center}>WebUrl</div>
          <div className={styles.Center} style={{ borderRight: 'none' }}>Visibility</div>
        </div>
        {this.state.ShownTeams.map(team => {
          team.Visibility = team.Visibility.substr(0, 1).toUpperCase() + team.Visibility.substr(1);
          var GroupEmailSplit = team.Mail.split("@");
          var Mail = GroupEmailSplit[0];
          return (
            <div className={styles.rowStyle}>
              <div className={styles.ToolTipName}>{team.Name}<span className={styles.ToolTip}>{team.Description}</span></div>
              <a className={styles.Center} href={`https://outlook.office365.com/mail/group/${this.Tenant}.com/${Mail.toLowerCase()}/email`}>
                <Icon className={iconClass} style={{ color: '#087CD7' }} iconName="OutlookLogo"></Icon></a>
              <a className={styles.Center} href={`https://${this.Tenant}.sharepoint.com/sites/${Mail}`}>
                <Icon className={iconClass} style={{ color: '#068B90' }} iconName="SharePointLogo"></Icon>
              </a>
              <a className={styles.Center} href={`https://outlook.office365.com/calendar/group/${this.Tenant}.com/${team.Name.replace(Replaceregex, '')}/view/week`}>
                <Icon className={iconClass} style={{ color: '#119AE2' }} iconName="Calendar"></Icon>
              </a>
              <div className={styles.Center}>
                {team.Planner === undefined ? <div></div> : <a href={team.Planner}>
                  <Icon className={iconClass} style={{ color: '#077D3F' }} iconName="ViewListTree"></Icon>
                </a>}
              </div>
              <a className={styles.Center} href={`${team.WebUrl}`}>
                <Icon className={iconClass} style={{ color: '#424AB5' }} iconName="TeamsLogo"></Icon>
              </a>
              <div className={styles.Center} style={{ borderRight: 'none' }}>{team.Visibility}</div>
            </div>
          );
        })}</div></div>;
  }
}