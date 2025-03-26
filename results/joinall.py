import os

def main():
    curr_dir = os.listdir()
    curr_dir = [f for f in curr_dir if f.endswith('.txt') and f != 'joined.txt']
    combined_file = os.curdir
    all_text = []
    skip = []

    curr_dir = sorted(curr_dir, key=lambda x: int(x.split('_')[1].split('.')[0].split('media')[1]))

    for file in curr_dir:
        fname = file.split("/")[-1]
        if fname.endswith('.txt'):
            with open(file, 'r') as f:
                line = f.readline()
                all_text.append(line)
    
    
    with open(f'{combined_file}/joined.txt', 'w') as out:
        slide_num = 1
        for t in all_text:
            while slide_num in skip:
                slide_num += 1
            out.write(f'{slide_num}.\n')
            out.write(t)
            out.write('\n\n')
            slide_num += 1


if __name__ == '__main__':
    main()
