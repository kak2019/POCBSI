import { spfi } from '@pnp/sp';
import { getSP } from '../../../../common/pnpjsConfig';

const REQUESTSCONST = { LIST_NAME: 'Nii Cases'};

const fetchById = async (arg: {
  Id: number;
}): Promise<Record<string, unknown> | string> => {
  const sp = spfi(getSP());
  const item = await sp.web.lists
    .getByTitle(REQUESTSCONST.LIST_NAME)
    .items.getById(arg.Id)()
    .catch((e) => e.message);
  return item;
};

const editRequest = async (arg: {
  request: Record<string, unknown>;
}): Promise<Record<string, unknown> | string> => {
  const { request } = arg;
  const sp = spfi(getSP());
  const list = sp.web.lists.getByTitle(REQUESTSCONST.LIST_NAME);
  await list.items
    .getById(+request.ID)
    .update(request)
    .catch((err) => err.message);
  const result = await fetchById({ Id: +request.ID });
  return result;
};

const addRequest  = async (arg: {
  request: Record<string, unknown>;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
}): Promise<Record<string, unknown> | string> => {
  const { request } = arg;
  const sp = spfi(getSP());
  const list = sp.web.lists.getByTitle(REQUESTSCONST.LIST_NAME);
  
  const result = await list.items.add(request).catch((err) => err.message);

  // const requestNew = result.data as Record<string, unknown>;
  // const titleStr = 'TAXI Request - ' + ('' + requestNew.ID).slice(-6);
  // const result2 = await editRequest({
  //   request: {
  //     ID: requestNew.ID,
  //     Title: titleStr,
  //   },
  // });

  return result;
};

const fetchUserGroups = async (arg: {
  userEmail: string;
}): Promise<string[]> => {
  try {
    const sp = spfi(getSP());
    const result: string[] = [];
    const user = await sp.web.ensureUser(arg.userEmail);
    const userId = user.data.Id;
    await sp.web.siteUsers
      .getById(userId)
      .groups()
      .then((response) =>
        response.map((o) => {
          result.push(o.Title);
        })
      );
    console.log(result);
    return result;
  } catch (err) {
    console.log(err);
    return Promise.reject("Error when fetch user's groups.");
  }
};

export { addRequest, editRequest, fetchById,fetchUserGroups };
