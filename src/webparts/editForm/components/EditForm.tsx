import * as React from 'react';
import { IEditFormProps } from './IEditFormProps';
import { IEditFormStates } from './IEditFormStates';
import { escape } from '@microsoft/sp-lodash-subset';
import { Container, Typography, FormControl, TextField, Select, MenuItem, Button } from '@material-ui/core';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export default class EditForm extends React.Component<IEditFormProps, IEditFormStates, any> {
  constructor(props) {
    super(props);

    // Set States (information managed within the component), When state changes, the component responds by re-rendering
    this.state = {
      taskID: '',
      task: '',
      description: '',
      priority: '',
      dueDate: null,
      taskExecutor: [],
      emailTaskExecutor: '',
      urlId: null,
      IsLoading: false
    };
  }

  componentDidMount = () => {
    // Start Loader
    this.setState({
      IsLoading: true
    });
    this.getPageId();
    console.log('2');
  }

  // For debugging (prints every update in state)
  // componentDidUpdate(prevProps: Readonly<IEditFormProps>, prevState: Readonly<any>, snapshot?: any): void {
  //   console.log('Updated', this.state);
  // }

  /**
   * Get item id from url
   */
  getPageId = () => {
    let url : any = new URL(window.location.href);
    let formId = url.searchParams.get("FormID");
    this.fetchData(formId);
  }
  
  /**
   * 
   * @param id 
   * @returns item data and set the state
   */
  fetchData = (id) => {
    let web = Web(this.props.webURL);
    return web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.getById(id)
      .get()
      .then(result => {
        this.setState({
          urlId: id,
          taskID: result.ID,
          task: result.Title,
          description: result.Description,
          priority: result.Priority,
          dueDate: result.Due_x0020_date.slice(0,10),
          taskExecutor: [result.Email_x0020_Task_x0020_Executor],
          emailTaskExecutor: result.Email_x0020_Task_x0020_Executor,
          IsLoading: false
        });
      }).catch(Err => {
        console.error(Err);
      });
  }

  /**
   * 
   * @param e 
   * Handeling input fields changes and set state
   */
  handleChange = (e: {target: {name: any; value: any; }; }) => {
    const newState = { [e.target.name]: e.target.value } as Pick<IEditFormStates, keyof IEditFormStates>;
    this.setState(newState);
  }

  /**
   * 
   * @param e 
   * Handeling select changes and set state
   */
  handleSelectChange = (e) => {
    this.setState({ priority: e.target.value });
  }

  /**
   * 
   * @param items 
   * Gets items(users) from share point and set state
   */
  getPeoplePickerItems = (items: any[]) => {
    console.log('Items:', items);
    this.setState({taskExecutor: items, emailTaskExecutor: items[0].secondaryText});
  }

  /**
   * Updates share point list
   */
  updateItem = () => {
    let web = Web(this.props.webURL);
    web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.getById(this.state.urlId).update({
      Title: this.state.task,
      Description: this.state.description,
      Priority: this.state.priority,
      Due_x0020_date: this.state.dueDate,
      Task_x0020_ExecutorId: this.state?.taskExecutor[0]?.id,
      Email_x0020_Task_x0020_Executor: this.state.emailTaskExecutor
    }).then(result => {
      alert("Item Updated Successfully");
      location.href = 'https://projects1.sharepoint.com/sites/Development/Alona/Lists/NewTasks/AllItems.aspx';
    }).catch (Err => {
      console.error(Err);
    })
  }


  public render(): React.ReactElement<IEditFormProps> {
    return (
      <Container maxWidth="sm">
        <Typography variant="h6" style={{ textAlign: 'center', marginTop: '10px', marginBottom: '10px' }}>Edit Task</Typography>
        { this.state.IsLoading ? 
          <div style={{ width: '100px', height: '50px', margin: '20px auto', textAlign: 'center', fontWeight: 'bold' }}>
            Loading...
          </div> :
          <div>
            <FormControl style={{ marginTop: '20px' }}>
              <TextField label="Task" name="task" value={this.state.task} variant="outlined" onChange={this.handleChange} style={{ marginTop: '13px', width: '500px' }}/>
              <TextField label="Task Description" name="description" value={this.state.description} variant="outlined" onChange={this.handleChange} multiline rows={3} style={{ marginTop: '13px', width: '500px' }}/>
              <Select
                label="Priority"
                name="priority"
                value={this.state.priority}
                variant="outlined" 
                onChange={(e) => {this.handleSelectChange(e)}}
                style={{ marginTop: '13px', width: '500px' }}
              >
                <MenuItem value="High">High</MenuItem>
                <MenuItem value="Medium">Medium</MenuItem>
                <MenuItem value="Low">Low</MenuItem>
              </Select>
              <TextField
                id="date"
                variant="outlined"
                label="Due Date"
                type="date"
                name="dueDate"
                value={this.state.dueDate}
                InputLabelProps={{
                  shrink: true,
                }}
                style={{ marginTop: '13px', width: '500px' }} 
                onChange={this.handleChange}
              />
              <PeoplePicker
                context={this.props.context as any}
                titleText="Task Executor"
                groupName={''}
                personSelectionLimit={1}
                required={false}
                showHiddenInUI={false}
                defaultSelectedUsers={this.state.taskExecutor}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                ensureUser={true}
                onChange={this.getPeoplePickerItems}
              />
            </FormControl>
            <div style={{ marginTop: '20px' }}>
              <Button style={{ width: '83px', marginRight: '5px'}} variant="outlined" onClick={this.updateItem} color="primary">
                Update
              </Button>
            </div>
          </div>
        }
      </Container>
    );
  }
}
