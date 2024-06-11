import * as React from 'react';
import { useState } from 'react';
import { PrimaryButton } from '@fluentui/react/lib/Button';


// 定义组件的 Props 类型
interface FileUploaderProps {
  onFileSelected: (file: File) => void; // 定义一个函数类型的 Prop，用于传递文件选择事件
}

const FileUploader: React.FC<FileUploaderProps> = ({ onFileSelected }) => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files ? event.target.files[0] : null;
    if (file) {
      setSelectedFile(file);
      onFileSelected(file);
    }
  };

  return (
    <div>
      <input
        type="file"
        style={{ display: 'none' }}
        id="fileInput"
        onChange={handleFileChange}
        accept="image/*" // 可以通过 accept 属性限制上传的文件类型
      />
      <label htmlFor="fileInput">
        <PrimaryButton
          // component="span"
          text={selectedFile ? `Upload: ${selectedFile.name}` : 'Upload File'}
          onClick={() => document.getElementById('fileInput')?.click()} // 当按钮被点击时触发文件选择
        />
      </label>
    </div>
  );
};

export default FileUploader;
