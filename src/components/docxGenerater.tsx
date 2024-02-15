import React, { useState } from 'react';
import Docxtemplater from 'docxtemplater';
import * as PizZip from 'pizzip';
import * as XLSX from 'xlsx';
import logo from '../assets/logo.svg';

interface TemplateGeneratorProps {}

const DocxGenerator: React.FC<TemplateGeneratorProps> = () => {
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [outputFiles, setOutputFiles] = useState<Blob[]>([]);
  const [outputFileNames, setOutputFileNames] = useState<string[]>([]);

  // 템플릿 파일 업로드 핸들러
  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setTemplateFile(e.target.files[0]);
    }
  };

  // 데이터 파일 업로드 핸들러
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setExcelFile(e.target.files[0]);
    }
  };

  // 결과 문서 생성
  const generateDocument = async () => {
    try {
      if (!templateFile || !excelFile) {
        throw new Error('템플릿 파일과 데이터 파일을 모두 선택해주세요.');
      }

      const templateData = await readFile(templateFile);
      const excelData = await readFile(excelFile);

      const wb = XLSX.read(excelData);
      const ws = wb.SheetNames;

      let data: any = [];
      ws.forEach((n) => {
        const ws = wb.Sheets[n];
        data = XLSX.utils.sheet_to_json(ws);
      });

      data.forEach((d: any) => {
        makeDocument(templateData, d);
      });
    } catch (error) {
      console.error(error);
    }
  };

  // DOCX 파일 생성 및 결과 저장
  const makeDocument = (template: any, data: any) => {
    const doc = new Docxtemplater();
    doc.loadZip(new PizZip(template));

    doc.setData(data);

    doc.render();
    const output = doc.getZip().generate({ type: 'blob' });

    setOutputFiles((prevFiles) => [...prevFiles, output]);
    setOutputFileNames((prevNames) => [...prevNames, data.name]);
    setTemplateFile(null);
    setExcelFile(null);
  };

  // 파일 읽기
  const readFile = (file: File): Promise<ArrayBuffer> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        if (e.target) {
          resolve(e.target.result as ArrayBuffer);
        }
      };

      reader.onerror = (error) => {
        reject(error);
      };

      reader.readAsArrayBuffer(file);
    });
  };

  //chorme multi file 다운로드를 위해 10개 마다 대기
  function pause(msec: number) {
    return new Promise(
        (resolve) => {
            setTimeout(resolve, msec || 1000);
        }
    );
  }

  // 결과 문서 다운로드
  const downloadDocument = () => {
    if (outputFiles.length > 0) {
      for(let i = 0; i< outputFiles.length; i ++){
        const link = document.createElement('a');
        link.href = URL.createObjectURL(outputFiles[i]);
        link.download = `${outputFileNames[i]}.docx`;
        
        if (i % 10 ==0) {
          await pause(1000).then(
            ()=>{
              link.click();
            }
          );
        } else{
          link.click();
        }
      }
    }
  };

  return (
    <>
      <div className="container">
        <div className="header">
          <div className="logo">
            <img src={logo} alt="logo"></img>
          </div>
          <div>A tool 🔧 that automatically Excel Data 🔡 into Docx Document 📑</div>
        </div>
        <div className="area-upload">
          <div className="upload-file">
            <span>📑 Templete File (.docx)</span>
            <input type="file" onChange={handleTemplateUpload} accept=".docx" />
          </div>
          <div className="area-icon">
            <span className="plus">+</span>
          </div>
          <div className="upload-file">
            <span>🔡 Data File (.excel)</span>
            <input type="file" onChange={handleExcelUpload} accept=".xlsx" />
          </div>
        </div>
        <div className="area-icon">
          <span className="equal"></span>
          <span className="equal"></span>
        </div>
        <div className="area-download">
          <button
            onClick={generateDocument}
            className={templateFile && excelFile && outputFiles.length === 0 ? 'active' : ''}
            disabled={!templateFile || !excelFile}
          >
            Make ✨
          </button>
          <div className="area-msg">
            {outputFiles.length > 0 && <div>{`"🎉 ${outputFiles.length} files maked. "`}</div>}
          </div>
          <button
            onClick={downloadDocument}
            className={outputFiles.length !== 0 ? 'active' : ''}
            disabled={outputFiles.length === 0}
          >
            All files Download 📚
          </button>
        </div>
      </div>
    </>
  );
};

export default DocxGenerator;
