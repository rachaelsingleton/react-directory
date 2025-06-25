import * as React from 'react';
import { useEffect, useState, useMemo, useCallback } from 'react';
import { MSGraphClient } from '@microsoft/sp-http';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { WebPartTitle } from '@pnp/spfx-controls-react';
import {
  Dropdown,
  Option,
  OptionOnSelectData,
  Overflow,
  OverflowItem,
  SearchBox,
  Field,
  Title2,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Tab,
  TabList,
  tokens,
  makeStyles,
  mergeClasses,
} from '@fluentui/react-components';
import { People48Filled } from '@fluentui/react-icons';

import { IDirectoryProps } from './IDirectoryProps';
import Paging from './Pagination/Paging';
import { OverflowAlphabetsMenu } from './OverflowAlphabetsMenu/OverflowAlphabetsMenu';
import styles from './Directory.module.scss';
import { PersonaCard } from './PersonaCard/PersonaCard';
import { IDirectoryState } from './IDirectoryState';
import * as strings from 'DirectoryWebPartStrings';
import { Shimmer } from './Shimmer/Shimmer';

const wrapStackTokens: IStackTokens = { childrenGap: 30 };
const useFluentStyles = makeStyles({
  alphabets: {
    backgroundColor: tokens.colorNeutralBackground2,
    overflow: 'hidden',
    padding: '5px',
    zIndex: 0,
    borderRadius: '5px',
  },
  horizontal: {
    height: 'fit-content',
    minWidth: '100%',
  },
  tabList: {
    justifyContent: 'center',
  },
});

const DirectoryHook: React.FC<IDirectoryProps> = (props) => {
  const {
    context,
    displayMode,
    title,
    updateProperty,
    searchFirstName,
    searchProps,
    pageSize: propPageSize,
    filterQuery,
    useSpaceBetween
  } = props;

  // 1) Parse the dynamic filterQuery which can contain multiple clauses
  interface IFilterClause { field: string; value: string; op?: 'and' | 'or' }

  const filterClauses = useMemo<IFilterClause[]>(() => {
    if (!filterQuery) return [];
    const parts = filterQuery.split(/\s+(AND|OR)\s+/i);
    const clauses: IFilterClause[] = [];
    let pendingOp: 'and' | 'or' | undefined = undefined;
    for (const part of parts) {
      if (!part.trim()) continue;
      if (/^and$/i.test(part)) {
        pendingOp = 'and';
        continue;
      }
      if (/^or$/i.test(part)) {
        pendingOp = 'or';
        continue;
      }
      const [raw, ...rest] = part.split(':');
      clauses.push({
        field: raw.trim().toLowerCase(),
        value: rest.join(':').trim().toLowerCase(),
        op: pendingOp,
      });
      pendingOp = undefined;
    }
    return clauses;
  }, [filterQuery]);

  // 2) Map alias → real Graph property names
  const graphFieldName = (f: string) => {
    switch (f.toLowerCase()) {
      case 'department':     return 'department';
      case 'officelocation': // alias "location:" or "officeLocation:"
      case 'location':       return 'officeLocation';
      case 'firstname':      return 'givenName';
      case 'lastname':       return 'surname';
      case 'workemail':      return 'mail';
      default:               return f;
    }
  };

  // 3) Apply filterQuery client-side
  const applyClientFilter = (users: any[]): any[] => {
    if (!filterClauses.length) return users;

    const matches = (u: any): boolean => {
      let result: boolean | undefined = undefined;
      filterClauses.forEach(clause => {
        const key = graphFieldName(clause.field);
        const match = String(u[key] || '')
          .toLowerCase()
          .trim() === clause.value;

        if (result === undefined) result = match;
        else if (clause.op === 'or') result = result || match;
        else result = result && match;
      });
      return result ?? true;
    };

    return users.filter(u => matches(u));
  };

  // 4) Build $filter for alpha & text searches (Graph-supported)
  const buildGraphFilter = (mode: 'initial' | 'alpha' | 'search'): string[] => {
    const clauses: string[] = [];

    // alpha (or initial) filter, skip when "All" selected
    if ((mode === 'initial' || mode === 'alpha') && alphaKey !== 'All') {
      const nameField = searchFirstName ? 'givenName' : 'surname';
      clauses.push(`startswith(${nameField},'${alphaKey}')`);
    }

    // text-search filter
    if (mode === 'search' && state.searchText.trim()) {
      const txt = state.searchText.trim().replace(/'/g, "''");
      const propsToSearch = searchProps
        ? searchProps.split(',').map(s => s.trim())
        : ['FirstName', 'LastName', 'WorkEmail', 'Department'];
      const sub = propsToSearch.map(p =>
        `startswith(${graphFieldName(p.toLowerCase())},'${txt}')`
      );
      if (sub.length) {
        clauses.push(`(${sub.join(' or ')})`);
      }
    }

    return clauses;
  };

  // 5) Component state
  const [az, setAz] = useState<string[]>([]);
  const [alphaKey, setAlphaKey] = useState<string>('All');
  const [state, setState] = useState<IDirectoryState>({
    users: [],
    isLoading: true,
    errorMessage: '',
    hasError: false,
    indexSelectedKey: 'All',
    searchString: 'LastName',
    searchText: '',
  });
  const [pagedItems, setPagedItems] = useState<any[]>([]);
  const [pageSize, setPageSize] = useState<number>(propPageSize ?? 10);
  const [currentPage, setCurrentPage] = useState<number>(1);

  useEffect(() => {
    setPageSize(propPageSize ?? 10);
  }, [propPageSize]);

  // 6) Paging helper
  const onPageUpdate = useCallback((pageNo?: number) => {
    const pg = pageNo ?? currentPage;
    const start = (pg - 1) * pageSize;
    const end = pg * pageSize;
    setPagedItems(state.users.slice(start, end));
    setCurrentPage(pg);
  }, [currentPage, pageSize, state.users]);

  // 7) Core fetch from Graph
  const fetchUsers = async (mode: 'initial' | 'alpha' | 'search') => {
    setState(s => ({ ...s, isLoading: true, hasError: false }));
    try {
      const client = await context.msGraphClientFactory.getClient('3');
      let req = client.api('/users').version('v1.0');
      const selectFields = [
        'id',
        'displayName',
        'jobTitle',
        'mail',
        'department',
        'officeLocation',
        'businessPhones'
      ].join(',');
      req = req.select(selectFields).top(999);

      // apply alpha/text filters server-side
      const odataClauses = buildGraphFilter(mode);
      if (odataClauses.length) {
        req = req.filter(odataClauses.join(' and '));
      }

      const resp: any = await req.get();
      let users: any[] = resp.value || [];

      // apply dynamic filterQuery client-side
      users = applyClientFilter(users);

      // determine which tab stays selected
      const selKey = mode === 'search' ? '0' : alphaKey;

      setState(s => ({
        ...s,
        users,
        isLoading: false,
        hasError: false,
        errorMessage: '',
        indexSelectedKey: selKey,
      }));
      setPagedItems(users.slice(0, pageSize));
      setCurrentPage(1);

    } catch (e: any) {
      setState(s => ({
        ...s,
        isLoading: false,
        hasError: true,
        errorMessage: e.message,
      }));
    }
  };

  // 8) Handlers
  const onAlphabetClick = (ev?: any) => {
    const key = ev?.target?.innerText ?? 'All';
    setAlphaKey(key);
    fetchUsers('alpha');
  };

  const doSearch = (text: string) => {
    setState(s => ({ ...s, searchText: text }));
    fetchUsers(text.trim() ? 'search' : 'initial');
  };

  const onSearchBoxChange = (_: any, data: { value: string }) => {
    doSearch(data.value);
  };

  const onSearchBoxKey = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') doSearch(state.searchText);
  }, [state.searchText]);

  // 9) Effects
  useEffect(() => {
    // build [All, A, B, … Z]
    const letters = Array.from({ length: 26 }, (_, i) =>
      String.fromCharCode(65 + i)
    );
    setAz(['All', ...letters]);

    fetchUsers('initial');
  }, [filterQuery]);

  // debounce text input
  useEffect(() => {
    const handle = setTimeout(() => doSearch(state.searchText), 300);
    return () => clearTimeout(handle);
  }, [state.searchText]);

  // update paging when users or pageSize change
  useEffect(() => {
    onPageUpdate(1);
  }, [state.users, pageSize]);

  // sort dropdown
  const onOptionSelect = (_: any, data: OptionOnSelectData) => {
    const field = data.optionValue as string;
    const sorted = [...state.users].sort((a, b) => {
      const aVal = String(a[field] || '').toLowerCase();
      const bVal = String(b[field] || '').toLowerCase();
      return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
    });
    setState(s => ({ ...s, users: sorted, searchString: field }));
  };

  const fluentStyles = useFluentStyles();
  const color = context.microsoftTeams ? 'white' : '';

  // 10) Render cards
  const directoryGrid = pagedItems.map((u, i) => (
    <PersonaCard
      key={i}
      context={context}
      profileProperties={{
        DisplayName: u.displayName,
        Title: u.jobTitle,
        PictureUrl: `/_layouts/15/userphoto.aspx?size=M&accountName=${u.mail}`,
        Email: u.mail,
        Department: u.department,
        WorkPhone: u.businessPhones?.[0],
        Location: u.officeLocation,
      }}
    />
  ));

  return (
    <div className={styles.directory}>
      <WebPartTitle
        displayMode={displayMode}
        title={title}
        updateProperty={updateProperty}
      />

      <div className={styles.searchBox}>
        <SearchBox
          placeholder={strings.SearchPlaceHolder}
          value={state.searchText}
          onKeyDown={onSearchBoxKey}
          onChange={onSearchBoxChange}
        />
        <div className={mergeClasses(fluentStyles.alphabets, fluentStyles.horizontal)}>
          <Overflow minimumVisible={2}>
            <TabList
              selectedValue={state.indexSelectedKey}
              onTabSelect={onAlphabetClick}
              className={fluentStyles.tabList}
            >
              {az.map(letter => (
                <OverflowItem key={letter} id={letter}>
                  <Tab value={letter}>{letter}</Tab>
                </OverflowItem>
              ))}
              <OverflowAlphabetsMenu tabs={az} onTabSelect={onAlphabetClick} />
            </TabList>
          </Overflow>
        </div>
      </div>

      {state.isLoading ? (
        <div style={{ marginTop: 10 }}><Shimmer /></div>
      ) : state.hasError ? (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>{state.errorMessage}</MessageBarTitle>
          </MessageBarBody>
        </MessageBar>
      ) : directoryGrid.length === 0 ? (
        <div className={styles.noUsers}>
          <People48Filled style={{ fontSize: 54, color }} />
          <Title2 style={{ marginLeft: 5, color }}>
            {strings.DirectoryMessage}
          </Title2>
        </div>
      ) : (
        <>
          <Paging
            totalItems={state.users.length}
            itemsCountPerPage={pageSize}
            currentPage={currentPage}
            onPageUpdate={onPageUpdate}
          />

          <div className={styles.dropDownSortBy}>
            <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
              <Field label={strings.DropDownPlaceLabelMessage}>
                <Dropdown
                  placeholder={strings.DropDownPlaceHolderMessage}
                  value={state.searchString}
                  onOptionSelect={onOptionSelect}
                >
                  {[
                    { value: 'displayName', text: 'Name' },
                    { value: 'jobTitle', text: 'Job Title' },
                    { value: 'department', text: 'Department' },
                    { value: 'officeLocation', text: 'Location' },
                  ].map(opt => (
                    <Option key={opt.value} value={opt.value}>
                      {opt.text}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
            </Stack>
          </div>

          <Stack
            horizontal
            wrap
            horizontalAlign={useSpaceBetween ? 'space-between' : 'center'}
            tokens={wrapStackTokens}
          >
            {directoryGrid}
          </Stack>

          <Paging
            totalItems={state.users.length}
            itemsCountPerPage={pageSize}
            currentPage={currentPage}
            onPageUpdate={onPageUpdate}
          />
        </>
      )}
    </div>
  );
};

export default DirectoryHook;
