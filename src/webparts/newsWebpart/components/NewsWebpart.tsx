import * as React from "react";
import styles from "./NewsWebpart.module.scss";
import { INewsWebpartProps } from "./INewsWebpartProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

interface INewsItem {
  Title: string;
  ShortDetail: string;
  NewsDetail: string;
  PublishedDate: any;
  ExpiryDate: any;
  NewsAuthor: {
    EMail: string;
    Title: string;
  };
  Author: {
    EMail: string;
    Title: string;
  };
  RedirectURL: {
    Url: string;
  };
  Department: any;
}

interface IAllNewsItems {
  AllNews: INewsItem[];
}

export default class NewsWebpart extends React.Component<
  INewsWebpartProps,
  IAllNewsItems
> {
  // Constructor
  constructor(props: INewsWebpartProps, state: IAllNewsItems) {
    super(props);
    this.state = {
      AllNews: [],
    };
  }
  componentDidMount() {
    this.getNewsData();
  }

  public getNewsData = () => {
    // Filter applied because we need to get Active links only from List
    let filterBy = `(PublishedDate le datetime'${new Date().toISOString()}') and (ExpiryDate ge datetime'${new Date().toISOString()}')`;

    let selectColumns = `*,NewsAuthor/Title,Author/Title`;
    let expandColumns = `NewsAuthor,Author`;

    let listURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=${selectColumns}&$filter=${filterBy}&$expand=${expandColumns}&$orderby=ID desc`;

    console.log(listURL);
    this.props.context.spHttpClient
      .get(listURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          this.setState({
            AllNews: responseJSON.value,
          });
          console.log(this.state.AllNews);
        });
      });
  };
  public render(): React.ReactElement<INewsWebpartProps> {
    return (
      <div className={styles["news-webpart"]}>
        <p>Hello there...</p>

        <div>
          {this.state.AllNews.map((news) => {
            return (
              <div>
                <p>{news.Title}</p>
                <p>{news.ShortDetail}</p>
                <p>{news.PublishedDate}</p>
                <p>{news.ExpiryDate}</p>
                <p>
                  {news.NewsAuthor == null
                    ? news.Author.Title
                    : news.NewsAuthor.Title}
                </p>
                <hr />
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
