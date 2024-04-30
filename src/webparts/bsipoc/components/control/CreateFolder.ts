import { spfi } from "@pnp/sp";
import { getSP } from "../../../../common/pnpjsConfig";
import "@pnp/sp/folders";
// 初始化SP PnP库，通常在应用程序的主入口文件中配置一次

// 创建文件夹的函数

export default async function createFolder(libraryUrl: string, folderName: string) {
    const sp = spfi(getSP());
    try {
        // 使用addUsingPath方法创建文件夹
        const folderAddResult = await sp.web.getFolderByServerRelativePath(libraryUrl).folders.addUsingPath(folderName);
        console.log(`Folder '${folderName}' created successfully!`);
        console.log(folderAddResult);
    } catch (error) {
        console.error("Error creating folder:", error);
    }
}


