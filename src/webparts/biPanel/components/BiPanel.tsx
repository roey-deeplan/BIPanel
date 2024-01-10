import * as React from 'react';
import styles from "./BiPanel.module.scss"
//import 'animate.css';
import "./BiPanelStyle.css";
import { IBiPanelProps } from './IBiPanelProps';
import { IBiPanelStates } from './IBiPanelStates';
import getSP from "../PnPjsConfig";

export default class BiPanel extends React.Component<IBiPanelProps, IBiPanelStates> {
  sp = getSP(this.props.context);

  constructor(props: IBiPanelProps) {
    super(props);
    // Set States (information managed within the component), When state changes, the component responds by re-rendering
    this.state = {
      IsLoading: false,
      ListOfReports: [],
      SelectedIframeUrl: ''
    };
  }

  componentDidMount() {
    // Start Loader
    this.setState({
      IsLoading: true
    });
    // Reset all the values in the panel
    this.ResetReport();
  }

  ResetReport = () => {
    // Get Current Web

    let CurrentIframe = '';
    // Get All Reports
    this.sp.web.lists.getById(this.props.listsData).items().then((Reports: any) => {
      // Set default report view
      if (Reports.length > 0) {
        CurrentIframe = Reports[0].BiIframeLink;
      }
      // Update data, stop loader and present the panel
      this.setState({
        IsLoading: false,
        ListOfReports: Reports,
        SelectedIframeUrl: CurrentIframe
      });
    }).catch((Err: Error) => {
      console.error(Err);
    });
  }

  // Change panel iframe
  ChangeIframe = (IframeLink: string) => {
    this.ScrollToTop();
    this.setState({
      SelectedIframeUrl: IframeLink
    });
  }

  // Scroll to the Top of the screen
  ScrollToTop = () => {
    window.scrollTo({
      top: 0,
      behavior: "smooth"
    });
  };

  ParseImage = (Image: string) => {
    const json = JSON.parse(Image)
    return json.serverUrl + json.serverRelativeUrl
  }

  public render(): React.ReactElement<IBiPanelProps> {
    return (
      <div className={styles.biPanel}>
        <div className="EONewFormContainer">
          {this.state.IsLoading ?
            <div className='SpinnerComp'>
              <div className="loading-screen">
                <div className="loader-wrap">
                  <span className="loader-animation"></span>
                  <div className="loading-text">
                    <span className='letter'>l</span>
                    <span className='letter'>o</span>
                    <span className='letter'>a</span>
                    <span className='letter'>d</span>
                    <span className='letter'>i</span>
                    <span className='letter'>n</span>
                    <span className='letter'>g</span>
                  </div>
                </div>
              </div>
            </div>
            :
            null}
          {!this.state.IsLoading ?
            <div className='MainPanel'>
              <section className='BiNav' style={{maxHeight: "750px", overflow: "auto", direction:"rtl" }}>
                <ul className='BiNavList'>

                  {this.state.ListOfReports.map(({ ID, Title, Image, BILink }) => (
                    <li className='BiNavItem' key={ID} onClick={() => this.ChangeIframe(BILink)}>
                      <div className='BiNavItemImg'>
                        {Image && <img src={this.ParseImage(Image)} alt={Title} />}
                      </div>
                      <div className='BiNavItemText'>{Title}</div>
                    </li>
                  ))}

                </ul>
              </section>
              <div className='BiReport'>
                <iframe src={this.state.SelectedIframeUrl} width='100%' height='720px' frameBorder="0" allowFullScreen={true} ></iframe>
              </div>
            </div>
            :
            null}
        </div>
      </div>
    );
  }
}