import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./Directory.module.scss";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { ISPServices } from "../../../SPServices/ISPServices";
import { IDirectoryState } from "./IDirectoryState";
import * as strings from "DirectoryWebPartStrings";
import { SearchBox } from "office-ui-fabric-react";
import { debounce } from "throttle-debounce";

import { IDirectoryProps } from './IDirectoryProps';
import Navbar from './Navbar/Navbar';
import { spservices } from '../../../SPServices/spservices';

const LogoSVG = (
  <svg width="42" height="38" viewBox="0 0 42 38" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M0 14.3725L16.0031 0H34.002C38.4186 0 42 3.24086 42 7.23749V23.5186L25.997 38V18.8559C25.997 16.3644 23.7603 14.3533 21.0071 14.3661H0.00708477L0 14.3725Z" fill="white" />
  </svg>
);

const DirectoryHook: React.FC<IDirectoryProps> = (props) => {
  const _services: ISPServices = new spservices(props.context);
  const [, setaz] = useState<string[]>([]);
  const [alphaKey, setalphaKey] = useState<string>('A');
  const [state, setState] = useState<IDirectoryState>({
    users: [],
    isLoading: true,
    errorMessage: "",
    hasError: false,
    indexSelectedKey: "A",
    searchString: "LastName",
    searchText: "",
    searchSuggestions: []
  });

  // Paging
  const [pagedItems, setPagedItems] = useState<any[]>([]);
  const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
  const [currentPage, setCurrentPage] = useState<number>(1);

  const _onPageUpdate = async (pageno?: number): Promise<void> => {
    const currentPge = (pageno) ? pageno : currentPage;
    const startItem = ((currentPge - 1) * pageSize);
    const endItem = currentPge * pageSize;
    const filItems = state.users.slice(startItem, endItem);
    setCurrentPage(currentPge);
    setPagedItems(filItems);
  };

  const _loadAlphabets = (): void => {
    const alphabets: string[] = [];
    for (let i = 65; i < 91; i++) {
      alphabets.push(
        String.fromCharCode(i)
      );
    }
    setaz(alphabets);
  };

  const searchUsers = async (searchString: string, searchFirstName: boolean): Promise<any> => {
    try {
      const users = await _services.searchUsers(searchString, searchFirstName);
      return users;
    } catch (error) {
      throw new Error(`Error while searching users: ${error.message}`);
    }
  };

  const _searchUsers = async (searchText: string): Promise<void> => {
    try {
      setState({ ...state, searchText: searchText, isLoading: true });
      const users = await searchUsers(searchText, props.searchFirstName);
      setState({
        ...state,
        searchText: searchText,
        indexSelectedKey: '0',
        users: users.PrimarySearchResults || [],
        isLoading: false,
        errorMessage: "",
        hasError: false
      });
      setalphaKey('0');
    } catch (err) {
      setState({ ...state, errorMessage: err.message, hasError: true });
    }
  };
  const _debouncesearchUsers = debounce(500, _searchUsers);

  const _searchBoxChanged = (newvalue: string): void => {
    setCurrentPage(1);
    _debouncesearchUsers(newvalue);
  };

  const _searchByAlphabets = async (initialSearch: boolean): Promise<void> => {
    setState({ ...state, isLoading: true, searchText: '' });
    let users = null;
    if (initialSearch) {
      if (props.searchFirstName)
        users = await _services.searchUsersNew('', `FirstName:a*`, false);
      else users = await _services.searchUsersNew('a', '', true);
    } else {
      if (props.searchFirstName)
        users = await _services.searchUsersNew('', `FirstName:${alphaKey}*`, false);
      else users = await _services.searchUsersNew(`${alphaKey}`, '', true);
    }
    setState({
      ...state,
      searchText: '',
      indexSelectedKey: initialSearch ? 'A' : state.indexSelectedKey,
      users:
        users && users.PrimarySearchResults
          ? users.PrimarySearchResults
          : null,
      isLoading: false,
      errorMessage: "",
      hasError: false
    });
  };
  

  useEffect(() => {
    if (props.pageSize !== undefined) {
      setPageSize(props.pageSize);
      if (state.users) {
        _onPageUpdate();
      }
    }
  }, [state.users, props.pageSize]);

  useEffect(() => {
    if (alphaKey && alphaKey.length > 0 && alphaKey !== "0") {
      _searchByAlphabets(false);
    }
  }, [alphaKey]);

  useEffect(() => {
    _loadAlphabets();
    _searchByAlphabets(true);
  }, [props]);

  return  (
    <div className={styles.directory}>
      <div className={styles.header}>
        <div className={styles.serchWrap}>
          <SearchBox
            placeholder={strings.SearchPlaceHolder}
            className={styles.searchTextBox}
            onSearch={_searchUsers}
            value={state.searchText}
            onChange={(ev, newVal) => {
              if (newVal !== undefined) {
                _searchBoxChanged(newVal);
              }
            }}
          />
          <div className={styles.logo}>{LogoSVG}</div>
          {state.searchText.length > 0 && pagedItems.length > 0 && (
            <div id="auto-suggest" >
              <ul className={styles.suggestions}>
                {pagedItems.map((user, index) => (
                  <li key={index} className={styles.suggestion}>
                    <PersonaCard
                      context={props.context}
                      key={"PersonaCard" + index}
                      profileProperties={{
                        DisplayName: user.PreferredName,
                        Title: user.JobTitle,
                        PictureUrl: user.PictureURL,
                        Email: user.WorkEmail,
                        Department: user.Department,
                        WorkPhone: user.WorkPhone,
                        Location: user.OfficeNumber
                          ? user.OfficeNumber
                          : user.BaseOfficeLocation
                      }}
                    />
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      </div>
      <div>
        <Navbar />
      </div>
    </div>
  );
};

export default DirectoryHook;
