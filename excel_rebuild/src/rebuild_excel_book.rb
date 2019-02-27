#encoding:utf-8
require 'zip' # gem install rubyzip
require "fileutils"



class ZipFileGenerator
  @work_folder_path
  # rubyzipのreadmeから拝借 フォルダの場合は再帰的に配下のファイルをzipファイルに追加 
  def write_entries(entries, path, zipfile)
    entries.each do |e|
      zipfile_path = path == '' ? e : File.join(path, e)
      disk_file_path = File.join(@work_folder_path, zipfile_path)
      puts "Deflating #{disk_file_path}"

      if File.directory? disk_file_path
        recursively_deflate_directory(disk_file_path, zipfile, zipfile_path)
      else
        put_into_archive(disk_file_path, zipfile, zipfile_path)
      end
    end
  end

  def recursively_deflate_directory(disk_file_path, zipfile, zipfile_path)
    zipfile.mkdir zipfile_path
    subdir = Dir.entries(disk_file_path) - %w(. ..)
    write_entries subdir, zipfile_path, zipfile
  end

  def put_into_archive(disk_file_path, zipfile, zipfile_path)
    zipfile.get_output_stream(zipfile_path) do |f|
      f.write(File.open(disk_file_path, 'rb').read)
    end
  end

  # 圧縮処理
  def compress (target_paths, zip_path, work_folder_path)
    @work_folder_path = work_folder_path
    File.unlink zip_path if File.file?(zip_path)
    Zip::File.open(zip_path, Zip::File::CREATE) do |zip_file|
      write_entries(target_paths, '', zip_file)
    end
  end
end

def rebuild
  root_path = File.absolute_path("..")
  work_folder_path = File.join(root_path, 'work')
  target_folder_path = File.join(root_path, 'target')
  result_folder_path = File.join(root_path, 'result')

  f = File.open(File.join(target_folder_path, 'xl', 'drawings', 'drawing1.xml'), mode = "r:utf-8:utf-8");
  deawinf_contents = f.read
  f.close

  # Encoding.default_external = 'utf-8'
  
  deawinf_contents = deawinf_contents.gsub(/(<xdr\:oneCellAnchor)/, "\r\n\\1").gsub(/(<xdr\:twoCellAnchor)/, "\r\n\\1").gsub(/(<\/xdr\:wsDr>)/,"\r\n\\1") # 改行
  # deawinf_contents = deawinf_contents.gsub(/(<xdr\:two)/, "\r\n\\1")
  lines = deawinf_contents.split("\r\n"); # 各行を配列へ
  # puts lines

  header = []
  footer = []
  header.push(lines.shift).push(lines.shift) # 先頭2行を取得
  footer.push(lines.pop) # 末尾1行を取得

  puts header
  puts footer

  cellAnchors = []
  lines.each_with_index do |line, i|
    cellAnchors.push(line)
    zip_path = File.join(result_folder_path, "rebuild_#{i}.xlsx")
    
    # workフォルダ内を削除
    empty_work_folder(work_folder_path)

    # targetの内容をworkにコピー
    copy2workdir(target_folder_path, work_folder_path)

    # drawing1.xmlの内容を1行ずつ増やして生成
    File.open(File.join(work_folder_path, 'xl', 'drawings', 'drawing1.xml') , "w") do |f| 
      f.puts((header + cellAnchors + footer).join("\r\n"));
    end

    # 出力
    zip_file_generator = ZipFileGenerator.new()
    puts zip_file_generator.compress((Dir.entries(work_folder_path) - %w(. ..)), zip_path, work_folder_path)

  end


  # empty_work_folder(work_folder_path)

end

def empty_work_folder(work_folder_path)
  (Dir.entries(work_folder_path) - %w(. ..)).each do |delete|
    FileUtils.rm_rf(File.join(work_folder_path, delete)) 
  end
end 

def copy2workdir(target_folder_path, work_folder_path)
  (Dir.entries(target_folder_path) - %w(. ..)).each do |entry|
    FileUtils.cp_r(File.join(target_folder_path, entry), work_folder_path)
  end
end

rebuild


